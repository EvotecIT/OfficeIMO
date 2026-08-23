using System;

namespace OfficeIMO.Drawing;

/// <summary>Portable metadata categories considered by image optimization.</summary>
[Flags]
public enum OfficeImageMetadataKinds {
    /// <summary>No metadata categories.</summary>
    None = 0,
    /// <summary>Exif metadata, excluding its separately reported orientation field.</summary>
    Exif = 1,
    /// <summary>Extensible Metadata Platform packets.</summary>
    Xmp = 2,
    /// <summary>Embedded ICC color profiles.</summary>
    Icc = 4,
    /// <summary>Embedded pixel orientation.</summary>
    Orientation = 8,
    /// <summary>Text comments.</summary>
    Comments = 16,
    /// <summary>Physical resolution or density.</summary>
    Resolution = 32,
    /// <summary>All defined portable metadata categories.</summary>
    All = Exif | Xmp | Icc | Orientation | Comments | Resolution
}

/// <summary>Policy applied to metadata when image bytes are rewritten.</summary>
public enum OfficeImageMetadataPolicy {
    /// <summary>Copy every safely supported source category and report any loss.</summary>
    Preserve,
    /// <summary>Remove portable source metadata from rewritten output.</summary>
    Strip,
    /// <summary>Copy only categories selected by <see cref="OfficeImageOptimizationRequest.MetadataSelection"/>.</summary>
    SelectiveCopy
}

/// <summary>Typed evidence describing metadata retained or lost by optimization.</summary>
public sealed class OfficeImageMetadataReport {
    internal OfficeImageMetadataReport(
        OfficeImageMetadataPolicy policy,
        OfficeImageMetadataKinds source,
        OfficeImageMetadataKinds requested,
        OfficeImageMetadataKinds preserved,
        OfficeImageMetadataKinds normalized) {
        Policy = policy;
        Source = source;
        Requested = requested;
        Preserved = preserved;
        Normalized = normalized;
    }

    /// <summary>Requested metadata policy.</summary>
    public OfficeImageMetadataPolicy Policy { get; }
    /// <summary>Categories discovered in the source container.</summary>
    public OfficeImageMetadataKinds Source { get; }
    /// <summary>Source categories selected for copying.</summary>
    public OfficeImageMetadataKinds Requested { get; }
    /// <summary>Selected categories present in the rewritten output.</summary>
    public OfficeImageMetadataKinds Preserved { get; }
    /// <summary>Categories whose semantics were retained after a required value rewrite.</summary>
    public OfficeImageMetadataKinds Normalized { get; }
    /// <summary>Selected categories that could not be retained.</summary>
    public OfficeImageMetadataKinds Lost => Requested & ~Preserved;
    /// <summary>Source categories deliberately removed by policy.</summary>
    public OfficeImageMetadataKinds Stripped => Source & ~Requested;
    /// <summary>Whether the rewrite lost any selected metadata category.</summary>
    public bool HasLoss => Lost != OfficeImageMetadataKinds.None;
}
