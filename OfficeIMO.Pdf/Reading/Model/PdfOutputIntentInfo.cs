namespace OfficeIMO.Pdf;

/// <summary>
/// Lightweight readback metadata for a catalog output intent.
/// </summary>
public sealed class PdfOutputIntentInfo {
    private readonly bool _hasDestinationOutputProfile;
    private readonly Lazy<PdfOutputIntentProfileMetadata?> _profileMetadata;

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
        bool hasDestinationOutputProfile,
        Func<PdfOutputIntentProfileMetadata?>? profileMetadataFactory) {
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
        _hasDestinationOutputProfile = hasDestinationOutputProfile;
        _profileMetadata = new Lazy<PdfOutputIntentProfileMetadata?>(
            profileMetadataFactory ?? (() => null),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication);
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

    /// <summary>True when /DestOutputProfile resolves to a profile stream.</summary>
    public bool HasDestinationOutputProfile => _hasDestinationOutputProfile;

    /// <summary>Object number of /DestOutputProfile when it is an indirect stream reference.</summary>
    public int? DestinationOutputProfileObjectNumber { get; }

    /// <summary>ICC profile stream /N component count, when present.</summary>
    public int? DestinationOutputProfileColorComponents { get; }

    /// <summary>ICC profile stream /Alternate color space, when present.</summary>
    public string? DestinationOutputProfileAlternateColorSpace { get; }

    /// <summary>ICC profile stream filter name or simple filter value, when present.</summary>
    public string? DestinationOutputProfileFilter { get; }

    /// <summary>ICC profile size in bytes after bounded stream decoding, when present.</summary>
    public int? DestinationOutputProfileSizeBytes => _profileMetadata.Value?.SizeBytes;

    /// <summary>Declared ICC profile size from the ICC header, when present and readable.</summary>
    public int? DestinationOutputProfileDeclaredSizeBytes => _profileMetadata.Value?.DeclaredSizeBytes;

    /// <summary>ICC profile color-space marker from the ICC header, for example RGB, GRAY, or CMYK.</summary>
    public string? DestinationOutputProfileColorSpace => _profileMetadata.Value?.ColorSpace;

    /// <summary>ICC profile device-class marker from the ICC header, for example scnr, mntr, or prtr.</summary>
    public string? DestinationOutputProfileDeviceClass => _profileMetadata.Value?.DeviceClass;

    /// <summary>True when the ICC header contains the acsp signature; false when a readable header is present without it.</summary>
    public bool? DestinationOutputProfileHasIccSignature => _profileMetadata.Value?.HasIccSignature;

    /// <summary>True when the ICC profile is parseable and exposes a supported PCS-to-device output transform.</summary>
    public bool? DestinationOutputProfileHasSupportedOutputTransform => _profileMetadata.Value?.HasSupportedOutputTransform;
}

internal sealed class PdfOutputIntentProfileMetadata {
    internal PdfOutputIntentProfileMetadata(
        int sizeBytes,
        int? declaredSizeBytes,
        string? colorSpace,
        string? deviceClass,
        bool? hasIccSignature,
        bool hasSupportedOutputTransform) {
        SizeBytes = sizeBytes;
        DeclaredSizeBytes = declaredSizeBytes;
        ColorSpace = colorSpace;
        DeviceClass = deviceClass;
        HasIccSignature = hasIccSignature;
        HasSupportedOutputTransform = hasSupportedOutputTransform;
    }

    internal int SizeBytes { get; }
    internal int? DeclaredSizeBytes { get; }
    internal string? ColorSpace { get; }
    internal string? DeviceClass { get; }
    internal bool? HasIccSignature { get; }
    internal bool HasSupportedOutputTransform { get; }
}
