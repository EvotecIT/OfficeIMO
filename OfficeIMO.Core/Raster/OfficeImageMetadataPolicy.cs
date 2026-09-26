using System;
using System.Collections.Generic;

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
public sealed class OfficeImageMetadataReport : IOfficeConversionReport {
    private readonly IReadOnlyList<OfficeConversionFidelityDiagnostic> _fidelityDiagnostics;

    internal OfficeImageMetadataReport(
        OfficeImageMetadataPolicy policy,
        OfficeImageMetadataKinds source,
        OfficeImageMetadataKinds requested,
        OfficeImageMetadataKinds preserved,
        OfficeImageMetadataKinds normalized,
        bool policyApplied) {
        Policy = policy;
        Source = source;
        Requested = requested;
        Preserved = preserved;
        Normalized = normalized;
        PolicyApplied = policyApplied;
        var fidelityDiagnostics = new List<OfficeConversionFidelityDiagnostic>();
        foreach (OfficeImageMetadataKinds kind in EnumerateKinds(Lost)) {
            fidelityDiagnostics.Add(new OfficeConversionFidelityDiagnostic(
                "IMAGE_METADATA_" + kind.ToString().ToUpperInvariant() + "_LOST",
                $"The requested {kind} image metadata could not be retained.",
                OfficeConversionLossKind.Omission,
                "OfficeIMO.Drawing.ImageMetadata",
                kind.ToString()));
        }
        _fidelityDiagnostics = fidelityDiagnostics.AsReadOnly();
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
    /// <summary>Whether the requested policy was applied to the returned encoded bytes.</summary>
    public bool PolicyApplied { get; }
    /// <summary>Selected categories that could not be retained.</summary>
    public OfficeImageMetadataKinds Lost => Requested & ~Preserved;
    /// <summary>Source categories deliberately removed by policy.</summary>
    public OfficeImageMetadataKinds Stripped => PolicyApplied
        ? Source & ~Requested
        : OfficeImageMetadataKinds.None;
    /// <summary>Whether the rewrite lost any selected metadata category.</summary>
    public bool HasLoss => Lost != OfficeImageMetadataKinds.None;

    /// <inheritdoc />
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics => _fidelityDiagnostics;

    /// <inheritdoc />
    public void RequireNoLoss() {
        if (HasLoss) throw new InvalidOperationException(
            "Image optimization lost requested metadata. Inspect FidelityDiagnostics for the omitted categories.");
    }

    private static IEnumerable<OfficeImageMetadataKinds> EnumerateKinds(OfficeImageMetadataKinds kinds) {
        foreach (OfficeImageMetadataKinds kind in new[] {
            OfficeImageMetadataKinds.Exif,
            OfficeImageMetadataKinds.Xmp,
            OfficeImageMetadataKinds.Icc,
            OfficeImageMetadataKinds.Orientation,
            OfficeImageMetadataKinds.Comments,
            OfficeImageMetadataKinds.Resolution
        }) {
            if ((kinds & kind) != 0) yield return kind;
        }
    }
}
