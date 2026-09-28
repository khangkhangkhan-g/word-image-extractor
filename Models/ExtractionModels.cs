using System.IO;
using System.Collections.ObjectModel;

namespace WordImageExtractor.Models;

public enum OutputImageFormat
{
    Original,
    Png,
    Jpeg,
    Tiff,
    Bmp
}

public enum OrganizationMode
{
    GroupByPage,
    Flat
}

public enum PaginationMode
{
    AccurateWord,
    Portable
}

public enum DocumentScope
{
    Body,
    Header,
    Footer
}

public sealed class ImageOccurrence
{
    public int Sequence { get; init; }
    public int ScopeSequence { get; init; }
    public DocumentScope Scope { get; init; }
    public string PartPath { get; init; } = string.Empty;
    public string RelationshipId { get; init; } = string.Empty;
    public string SourceAssetPath { get; init; } = string.Empty;
    public string OriginalFileName { get; init; } = string.Empty;
    public string OriginalExtension { get; init; } = string.Empty;
    public long SourceBytes { get; init; }
    public bool IsExternal { get; init; }
    public string? ExternalTarget { get; init; }
    public int? PageNumber { get; set; }
}

public sealed class ScanResult
{
    public string WorkingDocumentPath { get; init; } = string.Empty;
    public List<ImageOccurrence> Occurrences { get; init; } = [];
    public int BodyImageCount => Occurrences.Count(x => x.Scope == DocumentScope.Body);
    public int StaticImageCount => Occurrences.Count(x => x.Scope != DocumentScope.Body);
    public long TotalEmbeddedBytes => Occurrences.Where(x => !x.IsExternal).Sum(x => x.SourceBytes);
}

public sealed class PaginationResult
{
    public bool Success { get; init; }
    public int PageCount { get; init; }
    public int DetectedPictureCount { get; init; }
    public IReadOnlyList<int> Pages { get; init; } = Array.Empty<int>();
    public string? Warning { get; init; }
}

public sealed class ExtractionOptions
{
    public required string InputPath { get; init; }
    public required string OutputZipPath { get; init; }
    public OutputImageFormat OutputFormat { get; init; } = OutputImageFormat.Original;
    public OrganizationMode Organization { get; init; } = OrganizationMode.GroupByPage;
    public PaginationMode Pagination { get; init; } = PaginationMode.AccurateWord;
    public bool IncludeHeadersFooters { get; init; }
    public bool IncludeMetadata { get; init; } = true;
    public bool IncludeCsvManifest { get; init; } = true;
    public bool IncludeJsonManifest { get; init; } = true;
    public int JpegQuality { get; init; } = 95;
}

public sealed class ExtractionProgress
{
    public int Processed { get; init; }
    public int Total { get; init; }
    public string Stage { get; init; } = string.Empty;
    public string CurrentItem { get; init; } = string.Empty;
}

public sealed class ManifestEntry
{
    public int Sequence { get; set; }
    public string Scope { get; set; } = string.Empty;
    public int? Page { get; set; }
    public int PageIndex { get; set; }
    public string RelationshipId { get; set; } = string.Empty;
    public string SourceAsset { get; set; } = string.Empty;
    public string OriginalFileName { get; set; } = string.Empty;
    public long SourceBytes { get; set; }
    public string OutputPath { get; set; } = string.Empty;
    public string OutputFormat { get; set; } = string.Empty;
    public int? PixelWidth { get; set; }
    public int? PixelHeight { get; set; }
    public string Status { get; set; } = string.Empty;
    public string Notes { get; set; } = string.Empty;
}

public sealed class ExtractionSummary
{
    public int TotalInstances { get; set; }
    public int Written { get; set; }
    public int Failed { get; set; }
    public int ExternalSkipped { get; set; }
    public int? PageCount { get; set; }
    public bool UsedAccuratePagination { get; set; }
    public string OutputPath { get; set; } = string.Empty;
    public List<string> Warnings { get; } = [];
}

public sealed class PreparedDocument : IAsyncDisposable
{
    public required string OriginalPath { get; init; }
    public required string WorkingPath { get; init; }
    public bool IsTemporary { get; init; }

    public ValueTask DisposeAsync()
    {
        if (IsTemporary)
        {
            try { File.Delete(WorkingPath); } catch { }
        }
        return ValueTask.CompletedTask;
    }
}
