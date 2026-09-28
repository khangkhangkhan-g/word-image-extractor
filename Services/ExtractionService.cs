using System.IO;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using WordImageExtractor.Models;
using WordImageExtractor.Utilities;

namespace WordImageExtractor.Services;

public sealed class ExtractionService
{
    private const long MetadataDecodeThresholdBytes = 64L * 1024 * 1024;
    private const long ConversionMemorySafetyThresholdBytes = 256L * 1024 * 1024;

    private readonly DocxImageScanner _scanner = new();
    private readonly WordAutomationService _word = new();
    private readonly ImageConversionService _converter = new();

    public bool IsWordInstalled => _word.IsWordInstalled();

    public async Task<ScanResult> QuickScanAsync(string inputPath, bool includeHeadersFooters, CancellationToken cancellationToken)
    {
        var ext = Path.GetExtension(inputPath).ToLowerInvariant();
        if (ext == ".doc")
            return new ScanResult { WorkingDocumentPath = inputPath };
        return await _scanner.ScanAsync(inputPath, includeHeadersFooters, cancellationToken);
    }

    public Task<ExtractionSummary> ExtractAsync(
        ExtractionOptions options,
        IProgress<ExtractionProgress>? progress,
        CancellationToken cancellationToken)
        => Task.Run(() => ExtractCoreAsync(options, progress, cancellationToken), cancellationToken);

    private async Task<ExtractionSummary> ExtractCoreAsync(
        ExtractionOptions options,
        IProgress<ExtractionProgress>? progress,
        CancellationToken cancellationToken)
    {
        var summary = new ExtractionSummary { OutputPath = options.OutputZipPath };
        var partialPath = options.OutputZipPath + $".partial-{Guid.NewGuid():N}";

        try
        {
            progress?.Report(new ExtractionProgress { Stage = "Preparing document", Processed = 0, Total = 1 });
            await using var prepared = await _word.PrepareDocumentAsync(options.InputPath, cancellationToken);

            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(new ExtractionProgress { Stage = "Scanning images", Processed = 0, Total = 1 });
            var scan = await _scanner.ScanAsync(prepared.WorkingPath, options.IncludeHeadersFooters, cancellationToken);
            summary.TotalInstances = scan.Occurrences.Count;

            PaginationResult? pagination = null;
            if (options.Pagination == PaginationMode.AccurateWord && scan.BodyImageCount > 0)
            {
                progress?.Report(new ExtractionProgress { Stage = "Mapping Word pages", Processed = 0, Total = scan.BodyImageCount });
                pagination = await _word.MapBodyImagePagesAsync(prepared.WorkingPath, scan.BodyImageCount, cancellationToken);
                summary.PageCount = pagination.PageCount > 0 ? pagination.PageCount : null;
                summary.UsedAccuratePagination = pagination.Success;

                if (pagination.Success)
                {
                    var body = scan.Occurrences.Where(x => x.Scope == DocumentScope.Body).ToArray();
                    for (var i = 0; i < body.Length; i++)
                        body[i].PageNumber = pagination.Pages[i];
                }
                else if (!string.IsNullOrWhiteSpace(pagination.Warning))
                {
                    summary.Warnings.Add(pagination.Warning);
                }
            }
            else if (options.Pagination == PaginationMode.AccurateWord && scan.BodyImageCount == 0)
            {
                summary.Warnings.Add("No body image instances were found for page mapping.");
            }

            var outputDir = Path.GetDirectoryName(options.OutputZipPath);
            Directory.CreateDirectory(string.IsNullOrWhiteSpace(outputDir) ? Environment.CurrentDirectory : outputDir);
            if (File.Exists(partialPath)) File.Delete(partialPath);

            var manifest = new List<ManifestEntry>(scan.Occurrences.Count);
            using var sourceFs = new FileStream(prepared.WorkingPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 1024 * 1024, FileOptions.SequentialScan);
            using var sourceZip = new ZipArchive(sourceFs, ZipArchiveMode.Read, leaveOpen: false);
            using var outputFs = new FileStream(partialPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 1024 * 1024, FileOptions.SequentialScan);
            using var outputZip = new ZipArchive(outputFs, ZipArchiveMode.Create, leaveOpen: true);

            var perPageCounters = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var total = scan.Occurrences.Count;
            var processed = 0;

            foreach (var occurrence in scan.Occurrences.OrderBy(x => x.Sequence))
            {
                cancellationToken.ThrowIfCancellationRequested();
                processed++;
                progress?.Report(new ExtractionProgress
                {
                    Stage = "Extracting images",
                    Processed = processed,
                    Total = total,
                    CurrentItem = occurrence.OriginalFileName,
                });

                var manifestRow = new ManifestEntry
                {
                    Sequence = occurrence.Sequence,
                    Scope = occurrence.Scope.ToString().ToLowerInvariant(),
                    Page = occurrence.PageNumber,
                    RelationshipId = occurrence.RelationshipId,
                    SourceAsset = occurrence.SourceAssetPath,
                    OriginalFileName = occurrence.OriginalFileName,
                    SourceBytes = occurrence.SourceBytes,
                };
                manifest.Add(manifestRow);

                if (occurrence.IsExternal)
                {
                    manifestRow.Status = "external-skipped";
                    manifestRow.Notes = $"External linked image was not downloaded: {occurrence.ExternalTarget}";
                    summary.ExternalSkipped++;
                    continue;
                }

                var sourceEntry = sourceZip.GetEntry(occurrence.SourceAssetPath);
                if (sourceEntry is null)
                {
                    manifestRow.Status = "failed";
                    manifestRow.Notes = "Embedded image relationship points to a missing package entry.";
                    summary.Failed++;
                    continue;
                }

                try
                {
                    var groupKey = occurrence.Scope == DocumentScope.Body
                        ? (occurrence.PageNumber?.ToString(CultureInfo.InvariantCulture) ?? "unknown")
                        : occurrence.Scope.ToString().ToLowerInvariant();
                    perPageCounters.TryGetValue(groupKey, out var pageIndex);
                    pageIndex++;
                    perPageCounters[groupKey] = pageIndex;
                    manifestRow.PageIndex = pageIndex;

                    if (options.OutputFormat == OutputImageFormat.Original)
                    {
                        var ext = ImageConversionService.NormalizeExtension(occurrence.OriginalExtension);
                        var outputPath = BuildOutputPath(occurrence, pageIndex, ext, options.Organization);
                        manifestRow.OutputPath = outputPath;
                        manifestRow.OutputFormat = ext.TrimStart('.').ToLowerInvariant();

                        // For ordinary assets, metadata is read from one in-memory copy.
                        // For very large assets, preserve the streaming path and skip dimensions
                        // to avoid a large temporary allocation just for metadata.
                        if (options.IncludeMetadata && sourceEntry.Length <= MetadataDecodeThresholdBytes)
                        {
                            byte[] bytes;
                            using (var input = sourceEntry.Open())
                            using (var ms = new MemoryStream(capacity: sourceEntry.Length <= int.MaxValue ? (int)sourceEntry.Length : 0))
                            {
                                await input.CopyToAsync(ms, 1024 * 1024, cancellationToken);
                                bytes = ms.ToArray();
                            }
                            var dims = _converter.TryGetDimensions(bytes);
                            manifestRow.PixelWidth = dims.width;
                            manifestRow.PixelHeight = dims.height;

                            var outEntry = outputZip.CreateEntry(outputPath, CompressionLevel.NoCompression);
                            await using var outStream = outEntry.Open();
                            await outStream.WriteAsync(bytes, cancellationToken);
                        }
                        else
                        {
                            if (options.IncludeMetadata && sourceEntry.Length > MetadataDecodeThresholdBytes)
                                manifestRow.Notes = "Pixel dimensions skipped for this very large asset to protect memory; original bytes were preserved.";

                            var outEntry = outputZip.CreateEntry(outputPath, CompressionLevel.NoCompression);
                            await using var outStream = outEntry.Open();
                            await using var input = sourceEntry.Open();
                            await input.CopyToAsync(outStream, 1024 * 1024, cancellationToken);
                        }

                        manifestRow.Status = "ok";
                        summary.Written++;
                        continue;
                    }

                    if (sourceEntry.Length > ConversionMemorySafetyThresholdBytes)
                    {
                        // Raster conversion requires decoding the complete image. For extremely
                        // large single assets, preserve the original bytes instead of risking an
                        // out-of-memory crash. The document continues processing normally.
                        var originalExt = ImageConversionService.NormalizeExtension(occurrence.OriginalExtension);
                        var outputPath = BuildOutputPath(occurrence, pageIndex, originalExt, options.Organization);
                        manifestRow.OutputPath = outputPath;
                        manifestRow.OutputFormat = originalExt.TrimStart('.').ToLowerInvariant();
                        manifestRow.Status = "ok-original-fallback";
                        manifestRow.Notes = "Requested conversion was skipped for this asset because it exceeds the 256 MB per-image conversion safety threshold. Original bytes were preserved.";
                        var outEntry = outputZip.CreateEntry(outputPath, CompressionLevel.NoCompression);
                        await using var outStream = outEntry.Open();
                        await using var input = sourceEntry.Open();
                        await input.CopyToAsync(outStream, 1024 * 1024, cancellationToken);
                        summary.Written++;
                        summary.Warnings.Add($"{occurrence.OriginalFileName}: conversion skipped for a very large asset; original format was preserved.");
                        continue;
                    }

                    byte[] originalBytes;
                    using (var input = sourceEntry.Open())
                    using (var ms = new MemoryStream(capacity: sourceEntry.Length <= int.MaxValue ? (int)sourceEntry.Length : 0))
                    {
                        await input.CopyToAsync(ms, 1024 * 1024, cancellationToken);
                        originalBytes = ms.ToArray();
                    }

                    var converted = _converter.Convert(originalBytes, options.OutputFormat, options.JpegQuality, occurrence.OriginalExtension);
                    manifestRow.PixelWidth = converted.PixelWidth;
                    manifestRow.PixelHeight = converted.PixelHeight;

                    if (converted.Success && converted.Bytes is not null)
                    {
                        var outputPath = BuildOutputPath(occurrence, pageIndex, converted.Extension, options.Organization);
                        manifestRow.OutputPath = outputPath;
                        manifestRow.OutputFormat = converted.Extension.TrimStart('.').ToLowerInvariant();
                        var outEntry = outputZip.CreateEntry(outputPath, CompressionLevel.NoCompression);
                        await using var outStream = outEntry.Open();
                        await outStream.WriteAsync(converted.Bytes, cancellationToken);
                        manifestRow.Status = "ok";
                        summary.Written++;
                    }
                    else
                    {
                        // Critical safety rule: conversion failure never discards the source image.
                        // Preserve original bytes and record the fallback.
                        var originalExt = ImageConversionService.NormalizeExtension(occurrence.OriginalExtension);
                        var outputPath = BuildOutputPath(occurrence, pageIndex, originalExt, options.Organization);
                        manifestRow.OutputPath = outputPath;
                        manifestRow.OutputFormat = originalExt.TrimStart('.').ToLowerInvariant();
                        manifestRow.Status = "ok-original-fallback";
                        manifestRow.Notes = $"Requested conversion was unsupported or failed. Original bytes preserved. {converted.Error}";
                        var outEntry = outputZip.CreateEntry(outputPath, CompressionLevel.NoCompression);
                        await using var outStream = outEntry.Open();
                        await outStream.WriteAsync(originalBytes, cancellationToken);
                        summary.Written++;
                        summary.Warnings.Add($"{occurrence.OriginalFileName}: conversion failed; original format was preserved.");
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    manifestRow.Status = "failed";
                    manifestRow.Notes = ex.Message;
                    summary.Failed++;
                    summary.Warnings.Add($"{occurrence.OriginalFileName}: {ex.Message}");
                }
            }

            WriteReadme(outputZip, options, scan, summary, pagination);
            if (options.IncludeCsvManifest) WriteCsvManifest(outputZip, manifest);
            if (options.IncludeJsonManifest) WriteJsonManifest(outputZip, manifest, summary);

            outputZip.Dispose();
            outputFs.Flush(flushToDisk: true);
            outputFs.Dispose();

            cancellationToken.ThrowIfCancellationRequested();
            if (File.Exists(options.OutputZipPath)) File.Delete(options.OutputZipPath);
            File.Move(partialPath, options.OutputZipPath);
            return summary;
        }
        catch
        {
            try { if (File.Exists(partialPath)) File.Delete(partialPath); } catch { }
            throw;
        }
    }

    private static string BuildOutputPath(ImageOccurrence occurrence, int pageIndex, string extension, OrganizationMode organization)
    {
        var safeExt = ImageConversionService.NormalizeExtension(extension);
        var seq = occurrence.Sequence.ToString("D5", CultureInfo.InvariantCulture);

        if (occurrence.Scope == DocumentScope.Body)
        {
            var pageToken = occurrence.PageNumber is int p ? $"P{p:D4}" : "P_UNKNOWN";
            var fileName = $"{seq}_{pageToken}_I{pageIndex:D3}{safeExt}";
            if (organization == OrganizationMode.Flat)
                return $"images/{fileName}";

            var folder = occurrence.PageNumber is int page ? $"Page_{page:D4}" : "Page_Unknown";
            return $"images/{folder}/{fileName}";
        }

        var scopeName = occurrence.Scope == DocumentScope.Header ? "Headers" : "Footers";
        var staticFileName = $"{seq}_{occurrence.Scope.ToString().ToUpperInvariant()}_{pageIndex:D3}{safeExt}";
        return organization == OrganizationMode.Flat
            ? $"images/{staticFileName}"
            : $"images/Static/{scopeName}/{staticFileName}";
    }

    private static void WriteReadme(ZipArchive zip, ExtractionOptions options, ScanResult scan, ExtractionSummary summary, PaginationResult? pagination)
    {
        var text = new StringBuilder();
        text.AppendLine("WORD IMAGE EXTRACTOR - EXPORT PACKAGE");
        text.AppendLine("====================================");
        text.AppendLine();
        text.AppendLine($"Source: {Path.GetFileName(options.InputPath)}");
        text.AppendLine($"Image instances found: {scan.Occurrences.Count}");
        text.AppendLine($"Body images: {scan.BodyImageCount}");
        text.AppendLine($"Header/footer assets: {scan.StaticImageCount}");
        text.AppendLine($"Output format: {options.OutputFormat}");
        text.AppendLine($"Organization: {options.Organization}");
        text.AppendLine($"Pagination mode: {options.Pagination}");
        text.AppendLine($"Accurate page mapping used: {summary.UsedAccuratePagination}");
        if (summary.PageCount is int pages) text.AppendLine($"Document pages reported by Word: {pages}");
        if (pagination?.Warning is { Length: > 0 }) text.AppendLine($"Pagination note: {pagination.Warning}");
        text.AppendLine();
        text.AppendLine("QUALITY POLICY");
        text.AppendLine("- Original mode copies embedded image bytes without re-encoding.");
        text.AppendLine("- Conversion keeps original pixel dimensions and never intentionally upscales.");
        text.AppendLine("- If conversion is unsupported or fails, the original image is preserved instead of being dropped.");
        text.AppendLine();
        text.AppendLine("PAGE MAPPING POLICY");
        text.AppendLine("- Accurate mode uses Microsoft Word's own pagination engine.");
        text.AppendLine("- Page numbers are applied only when Word picture-object count exactly matches body image-instance count.");
        text.AppendLine("- On mismatch, page assignment is withheld rather than guessed; document order is still preserved.");
        text.AppendLine();
        text.AppendLine("Copyright © 2026 Nguyen Khang. All Rights Reserved.");

        WriteTextEntry(zip, "README_FIRST.txt", text.ToString());
    }

    private static void WriteCsvManifest(ZipArchive zip, IReadOnlyList<ManifestEntry> manifest)
    {
        var sb = new StringBuilder();
        sb.AppendLine("Sequence,Scope,Page,PageIndex,RelationshipId,SourceAsset,OriginalFileName,SourceBytes,OutputPath,OutputFormat,PixelWidth,PixelHeight,Status,Notes");
        foreach (var m in manifest)
        {
            var values = new[]
            {
                m.Sequence.ToString(CultureInfo.InvariantCulture), m.Scope, m.Page?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                m.PageIndex.ToString(CultureInfo.InvariantCulture), m.RelationshipId, m.SourceAsset, m.OriginalFileName,
                m.SourceBytes.ToString(CultureInfo.InvariantCulture), m.OutputPath, m.OutputFormat,
                m.PixelWidth?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                m.PixelHeight?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                m.Status, m.Notes
            };
            sb.AppendLine(string.Join(',', values.Select(Csv.Escape)));
        }
        WriteTextEntry(zip, "metadata/manifest.csv", sb.ToString());
    }

    private static void WriteJsonManifest(ZipArchive zip, IReadOnlyList<ManifestEntry> manifest, ExtractionSummary summary)
    {
        var payload = new
        {
            generated_at = DateTimeOffset.Now,
            summary = new
            {
                total_instances = summary.TotalInstances,
                written = summary.Written,
                failed = summary.Failed,
                external_skipped = summary.ExternalSkipped,
                page_count = summary.PageCount,
                accurate_pagination = summary.UsedAccuratePagination,
                warnings = summary.Warnings,
            },
            images = manifest,
        };
        var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions { WriteIndented = true });
        WriteTextEntry(zip, "metadata/manifest.json", json);
    }

    private static void WriteTextEntry(ZipArchive zip, string path, string contents)
    {
        var entry = zip.CreateEntry(path, CompressionLevel.Optimal);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        writer.Write(contents);
    }
}
