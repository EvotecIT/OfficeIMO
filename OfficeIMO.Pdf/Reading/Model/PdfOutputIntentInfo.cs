namespace OfficeIMO.Pdf;

/// <summary>
/// Lightweight readback metadata for a catalog output intent.
/// </summary>
public sealed class PdfOutputIntentInfo {
    internal PdfOutputIntentInfo(
        int? objectNumber,
        string? subtype,
        string? outputConditionIdentifier,
        string? outputCondition,
        string? registryName,
        string? info,
        int? destinationOutputProfileObjectNumber,
        int? destinationOutputProfileColorComponents,
        string? destinationOutputProfileAlternateColorSpace,
        string? destinationOutputProfileFilter,
        int? destinationOutputProfileSizeBytes,
        int? destinationOutputProfileDeclaredSizeBytes,
        string? destinationOutputProfileColorSpace,
        bool? destinationOutputProfileHasIccSignature) {
        ObjectNumber = objectNumber;
        Subtype = subtype;
        OutputConditionIdentifier = outputConditionIdentifier;
        OutputCondition = outputCondition;
        RegistryName = registryName;
        Info = info;
        DestinationOutputProfileObjectNumber = destinationOutputProfileObjectNumber;
        DestinationOutputProfileColorComponents = destinationOutputProfileColorComponents;
        DestinationOutputProfileAlternateColorSpace = destinationOutputProfileAlternateColorSpace;
        DestinationOutputProfileFilter = destinationOutputProfileFilter;
        DestinationOutputProfileSizeBytes = destinationOutputProfileSizeBytes;
        DestinationOutputProfileDeclaredSizeBytes = destinationOutputProfileDeclaredSizeBytes;
        DestinationOutputProfileColorSpace = destinationOutputProfileColorSpace;
        DestinationOutputProfileHasIccSignature = destinationOutputProfileHasIccSignature;
    }

    /// <summary>Output intent object number when the output intent is indirect.</summary>
    public int? ObjectNumber { get; }

    /// <summary>Output intent /S subtype, for example GTS_PDFA1.</summary>
    public string? Subtype { get; }

    /// <summary>Output condition identifier from /OutputConditionIdentifier.</summary>
    public string? OutputConditionIdentifier { get; }

    /// <summary>Human-readable /OutputCondition value, when present.</summary>
    public string? OutputCondition { get; }

    /// <summary>Registry name from /RegistryName, when present.</summary>
    public string? RegistryName { get; }

    /// <summary>Human-readable /Info value, when present.</summary>
    public string? Info { get; }

    /// <summary>True when /DestOutputProfile metadata was readable.</summary>
    public bool HasDestinationOutputProfile => DestinationOutputProfileObjectNumber.HasValue || DestinationOutputProfileSizeBytes.HasValue;

    /// <summary>Object number of /DestOutputProfile when it is an indirect stream reference.</summary>
    public int? DestinationOutputProfileObjectNumber { get; }

    /// <summary>ICC profile stream /N component count, when present.</summary>
    public int? DestinationOutputProfileColorComponents { get; }

    /// <summary>ICC profile stream /Alternate color space, when present.</summary>
    public string? DestinationOutputProfileAlternateColorSpace { get; }

    /// <summary>ICC profile stream filter name or simple filter value, when present.</summary>
    public string? DestinationOutputProfileFilter { get; }

    /// <summary>ICC profile stream size in bytes after PDF stream parsing, when present.</summary>
    public int? DestinationOutputProfileSizeBytes { get; }

    /// <summary>Declared ICC profile size from the ICC header, when present and readable.</summary>
    public int? DestinationOutputProfileDeclaredSizeBytes { get; }

    /// <summary>ICC profile color-space marker from the ICC header, for example RGB, GRAY, or CMYK.</summary>
    public string? DestinationOutputProfileColorSpace { get; }

    /// <summary>True when the ICC header contains the acsp signature; false when a readable header is present without it.</summary>
    public bool? DestinationOutputProfileHasIccSignature { get; }
}
