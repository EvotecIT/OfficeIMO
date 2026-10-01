using OfficeIMO.Drawing;

namespace OfficeIMO.Word;

/// <summary>Outcome for one unique embedded media part.</summary>
public enum WordImageOptimizationStatus {
    /// <summary>Candidate image bytes are smaller or replacement was explicitly permitted.</summary>
    Optimized,
    /// <summary>The original image already meets the requested policy.</summary>
    AlreadySuitable,
    /// <summary>The candidate was not smaller.</summary>
    OriginalWasSmaller,
    /// <summary>The managed engine does not rewrite this format.</summary>
    UnsupportedFormat,
    /// <summary>The original could not be decoded safely.</summary>
    DecodeFailed,
    /// <summary>At least one usage has geometry that cannot be resolved safely.</summary>
    UnknownPlacement,
    /// <summary>No XML image reference was found; the part is preserved.</summary>
    Unreferenced,
    /// <summary>The candidate would lose metadata without explicit permission.</summary>
    MetadataLoss
}

/// <summary>Immutable encoded-media evidence for one unique package image.</summary>
public sealed class WordImageOptimizationItem {
    internal WordImageOptimizationItem(string partUri, int references, WordImageOptimizationStatus status,
        long originalBytes, long finalBytes, OfficeImageInfo original, OfficeImageInfo final,
        OfficeImageMetadataReport? metadata = null) {
        PartUri = partUri; ReferenceCount = references; Status = status;
        OriginalBytes = originalBytes; FinalBytes = finalBytes;
        Original = original; Final = final; Metadata = metadata;
    }
    /// <summary>Original package-relative media URI. BMP/GIF conversion creates a PNG carrier with the same drawing relationship IDs.</summary>
    public string PartUri { get; }
    /// <summary>Number of XML relationship references to this media part.</summary>
    public int ReferenceCount { get; }
    /// <summary>Optimization or preservation outcome.</summary>
    public WordImageOptimizationStatus Status { get; }
    /// <summary>Encoded bytes before optimization.</summary>
    public long OriginalBytes { get; }
    /// <summary>Encoded bytes after the selected policy; original size for preserved parts.</summary>
    public long FinalBytes { get; }
    /// <summary>Signed encoded-media savings.</summary>
    public long BytesSaved => OriginalBytes - FinalBytes;
    /// <summary>Original format and pixel dimensions.</summary>
    public OfficeImageInfo Original { get; }
    /// <summary>Candidate format and pixel dimensions, or original properties for preserved parts.</summary>
    public OfficeImageInfo Final { get; }
    /// <summary>Metadata evidence when an encoding candidate was evaluated.</summary>
    public OfficeImageMetadataReport? Metadata { get; }
}

/// <summary>Exact media-byte evidence. Whole-file savings are measured separately after saving.</summary>
public sealed class WordImageOptimizationReport {
    internal WordImageOptimizationReport(IEnumerable<WordImageOptimizationItem> images, int externalReferences, bool applied) {
        Images = Array.AsReadOnly(images.ToArray());
        ExternalReferenceCount = externalReferences; Applied = applied;
    }
    /// <summary>One entry per unique embedded image part, including preserved parts.</summary>
    public IReadOnlyList<WordImageOptimizationItem> Images { get; }
    /// <summary>Whether candidates were applied to the document. False for analysis.</summary>
    public bool Applied { get; }
    /// <summary>Externally linked image references skipped without fetching their targets.</summary>
    public int ExternalReferenceCount { get; }
    /// <summary>Unique embedded image count.</summary>
    public int ImageCount => Images.Count;
    /// <summary>Number of applicable candidates; proposed changes when Applied is false.</summary>
    public int OptimizedCount => Images.Count(image => image.Status == WordImageOptimizationStatus.Optimized);
    /// <summary>Signed encoded-media savings, or exact candidate savings during analysis.</summary>
    public long BytesSaved => Images.Sum(image => image.BytesSaved);
}
