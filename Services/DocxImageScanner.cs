using System.IO;
using System.IO.Compression;
using System.Xml.Linq;
using WordImageExtractor.Models;

namespace WordImageExtractor.Services;

public sealed class DocxImageScanner
{
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace V = "urn:schemas-microsoft-com:vml";
    private static readonly XNamespace Asvg = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";

    public Task<ScanResult> ScanAsync(string docxPath, bool includeHeadersFooters, CancellationToken cancellationToken)
        => Task.Run(() => Scan(docxPath, includeHeadersFooters, cancellationToken), cancellationToken);

    private static ScanResult Scan(string docxPath, bool includeHeadersFooters, CancellationToken cancellationToken)
    {
        using var fs = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 1024 * 1024, FileOptions.SequentialScan);
        using var archive = new ZipArchive(fs, ZipArchiveMode.Read, leaveOpen: false);

        var parts = new List<(string path, DocumentScope scope)> { ("word/document.xml", DocumentScope.Body) };
        if (includeHeadersFooters)
        {
            parts.AddRange(archive.Entries
                .Where(e => e.FullName.StartsWith("word/header", StringComparison.OrdinalIgnoreCase) && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                .OrderBy(e => e.FullName, StringComparer.OrdinalIgnoreCase)
                .Select(e => (e.FullName, DocumentScope.Header)));
            parts.AddRange(archive.Entries
                .Where(e => e.FullName.StartsWith("word/footer", StringComparison.OrdinalIgnoreCase) && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                .OrderBy(e => e.FullName, StringComparer.OrdinalIgnoreCase)
                .Select(e => (e.FullName, DocumentScope.Footer)));
        }

        var occurrences = new List<ImageOccurrence>();
        var globalSequence = 0;
        var scopeCounters = new Dictionary<DocumentScope, int>
        {
            [DocumentScope.Body] = 0,
            [DocumentScope.Header] = 0,
            [DocumentScope.Footer] = 0,
        };

        foreach (var (partPath, scope) in parts)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var partEntry = archive.GetEntry(partPath);
            if (partEntry is null) continue;

            XDocument xml;
            using (var stream = partEntry.Open())
                xml = XDocument.Load(stream, LoadOptions.None);

            var rels = ReadRelationships(archive, partPath);

            foreach (var element in xml.Descendants())
            {
                cancellationToken.ThrowIfCancellationRequested();

                // In mc:AlternateContent, Word commonly stores both a modern DrawingML
                // representation and a VML fallback for the same picture. Ignore the
                // fallback branch so the same visible image is not exported twice.
                if (element.Ancestors(Mc + "Fallback").Any())
                    continue;

                string? relId = null;
                if (element.Name == A + "blip")
                {
                    // Modern Office SVG pictures often contain a raster fallback on a:blip
                    // plus the true SVG relationship in asvg:svgBlip. Prefer the SVG asset
                    // so one visible picture produces one highest-fidelity export instance.
                    var svgBlip = element.Descendants(Asvg + "svgBlip").FirstOrDefault();
                    relId = (string?)svgBlip?.Attribute(R + "embed")
                        ?? (string?)element.Attribute(R + "embed")
                        ?? (string?)element.Attribute(R + "link");
                }
                else if (element.Name == V + "imagedata")
                    relId = (string?)element.Attribute(R + "id");

                if (string.IsNullOrWhiteSpace(relId))
                    continue;

                if (!rels.TryGetValue(relId, out var relation))
                    continue;

                // Only image relationships are valid extraction candidates.
                if (!relation.Type.EndsWith("/image", StringComparison.OrdinalIgnoreCase))
                    continue;

                globalSequence++;
                scopeCounters[scope]++;

                if (relation.IsExternal)
                {
                    occurrences.Add(new ImageOccurrence
                    {
                        Sequence = globalSequence,
                        ScopeSequence = scopeCounters[scope],
                        Scope = scope,
                        PartPath = partPath,
                        RelationshipId = relId,
                        SourceAssetPath = string.Empty,
                        OriginalFileName = SafeFileNameFromTarget(relation.Target, $"external_{globalSequence:D5}"),
                        OriginalExtension = Path.GetExtension(relation.Target),
                        IsExternal = true,
                        ExternalTarget = relation.Target,
                    });
                    continue;
                }

                var resolvedPath = ResolveTargetPath(partPath, relation.Target);
                var mediaEntry = archive.GetEntry(resolvedPath);
                if (mediaEntry is null)
                {
                    // Preserve the unresolved relationship in the manifest instead of
                    // silently dropping it. ExtractionService will mark it failed.
                    occurrences.Add(new ImageOccurrence
                    {
                        Sequence = globalSequence,
                        ScopeSequence = scopeCounters[scope],
                        Scope = scope,
                        PartPath = partPath,
                        RelationshipId = relId,
                        SourceAssetPath = resolvedPath,
                        OriginalFileName = Path.GetFileName(resolvedPath),
                        OriginalExtension = Path.GetExtension(resolvedPath),
                        SourceBytes = 0,
                    });
                    continue;
                }

                occurrences.Add(new ImageOccurrence
                {
                    Sequence = globalSequence,
                    ScopeSequence = scopeCounters[scope],
                    Scope = scope,
                    PartPath = partPath,
                    RelationshipId = relId,
                    SourceAssetPath = resolvedPath,
                    OriginalFileName = Path.GetFileName(resolvedPath),
                    OriginalExtension = Path.GetExtension(resolvedPath),
                    SourceBytes = mediaEntry.Length,
                });
            }
        }

        return new ScanResult
        {
            WorkingDocumentPath = docxPath,
            Occurrences = occurrences,
        };
    }

    private static Dictionary<string, RelationshipInfo> ReadRelationships(ZipArchive archive, string partPath)
    {
        var relsPath = RelationshipPartPath(partPath);
        var relEntry = archive.GetEntry(relsPath);
        var map = new Dictionary<string, RelationshipInfo>(StringComparer.Ordinal);
        if (relEntry is null) return map;

        XDocument relXml;
        using (var stream = relEntry.Open())
            relXml = XDocument.Load(stream, LoadOptions.None);

        foreach (var rel in relXml.Root?.Elements(Rel + "Relationship") ?? Enumerable.Empty<XElement>())
        {
            var id = (string?)rel.Attribute("Id");
            var type = (string?)rel.Attribute("Type") ?? string.Empty;
            var target = (string?)rel.Attribute("Target") ?? string.Empty;
            var targetMode = (string?)rel.Attribute("TargetMode");
            if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(target))
                continue;

            map[id] = new RelationshipInfo(type, target, string.Equals(targetMode, "External", StringComparison.OrdinalIgnoreCase));
        }
        return map;
    }

    private static string RelationshipPartPath(string partPath)
    {
        var slash = partPath.LastIndexOf('/');
        var dir = slash >= 0 ? partPath[..slash] : string.Empty;
        var file = slash >= 0 ? partPath[(slash + 1)..] : partPath;
        return string.IsNullOrEmpty(dir) ? $"_rels/{file}.rels" : $"{dir}/_rels/{file}.rels";
    }

    private static string ResolveTargetPath(string sourcePart, string target)
    {
        target = Uri.UnescapeDataString(target);
        if (target.StartsWith('/'))
            return target.TrimStart('/');

        var slash = sourcePart.LastIndexOf('/');
        var sourceDir = slash >= 0 ? sourcePart[..slash] : string.Empty;
        var components = new List<string>();
        if (!string.IsNullOrEmpty(sourceDir))
            components.AddRange(sourceDir.Split('/', StringSplitOptions.RemoveEmptyEntries));

        foreach (var piece in target.Replace('\\', '/').Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (piece == ".") continue;
            if (piece == "..")
            {
                if (components.Count > 0) components.RemoveAt(components.Count - 1);
                continue;
            }
            components.Add(piece);
        }
        return string.Join('/', components);
    }

    private static string SafeFileNameFromTarget(string target, string fallback)
    {
        try
        {
            var uri = new Uri(target, UriKind.RelativeOrAbsolute);
            var name = uri.IsAbsoluteUri ? Path.GetFileName(uri.LocalPath) : Path.GetFileName(target);
            return string.IsNullOrWhiteSpace(name) ? fallback : name;
        }
        catch
        {
            return fallback;
        }
    }

    private sealed record RelationshipInfo(string Type, string Target, bool IsExternal);
}
