namespace OfficeIMO.Html;

/// <summary>Describes one built-in HTML target, its public APIs, and feature-level contract.</summary>
public sealed class HtmlTargetCapabilityContract {
    private readonly IReadOnlyList<string> _profiles;

    /// <summary>Creates a target contract from independently classified conversion routes.</summary>
    public HtmlTargetCapabilityContract(
        HtmlConversionTarget target,
        string packageName,
        string artifactName,
        HtmlToTargetCapabilityContract htmlToTarget,
        TargetToHtmlCapabilityContract? targetToHtml = null) {
        Target = target;
        PackageName = Required(packageName, nameof(packageName));
        ArtifactName = Required(artifactName, nameof(artifactName));
        HtmlToTarget = htmlToTarget ?? throw new ArgumentNullException(nameof(htmlToTarget));
        TargetToHtml = targetToHtml;
        _profiles = HtmlToTarget.Profiles.Concat(TargetToHtml?.Profiles ?? Array.Empty<string>())
            .Distinct(StringComparer.Ordinal)
            .ToList()
            .AsReadOnly();
    }

    /// <summary>
    /// Creates a compatibility target contract. New catalogs should use the directional constructor
    /// so import and reverse-export profiles and feature outcomes cannot be conflated.
    /// </summary>
    [Obsolete("Use the constructor that accepts HtmlToTargetCapabilityContract and TargetToHtmlCapabilityContract.")]
    public HtmlTargetCapabilityContract(
        HtmlConversionTarget target,
        string packageName,
        string artifactName,
        string importEntryPoint,
        string importResultContract,
        string? exportEntryPoint,
        string? exportResultContract,
        string ioAndAsyncBoundary,
        IEnumerable<string> profiles,
        IEnumerable<HtmlSemanticFeature> supported,
        IEnumerable<HtmlSemanticFeature> approximated,
        IEnumerable<HtmlSemanticFeature> unsupported)
        : this(target, packageName, artifactName,
            new HtmlToTargetCapabilityContract(importEntryPoint, importResultContract, ioAndAsyncBoundary,
                "HtmlConversionReport diagnostics on the route result.",
                profiles, supported, approximated, unsupported),
            Optional(exportEntryPoint) == null ? null : new TargetToHtmlCapabilityContract(
                exportEntryPoint!, Required(exportResultContract!, nameof(exportResultContract)), ioAndAsyncBoundary,
                "HtmlConversionReport diagnostics on the route result.",
                profiles, supported, approximated, unsupported)) { }

    /// <summary>Target identifier used by preflight and adapter selection.</summary>
    public HtmlConversionTarget Target { get; }
    /// <summary>Package that owns the thin target projection.</summary>
    public string PackageName { get; }
    /// <summary>Native or rendered artifact produced by the target.</summary>
    public string ArtifactName { get; }
    /// <summary>HTML-to-target route, including its own profiles and feature partition.</summary>
    public HtmlToTargetCapabilityContract HtmlToTarget { get; }
    /// <summary>Target-to-HTML route, or <see langword="null"/> when unavailable.</summary>
    public TargetToHtmlCapabilityContract? TargetToHtml { get; }
    /// <summary>Primary public HTML import entry point.</summary>
    public string ImportEntryPoint => HtmlToTarget.EntryPoint;
    /// <summary>Public result or evidence contract returned by the import path.</summary>
    public string ImportResultContract => HtmlToTarget.ResultContract;
    /// <summary>Primary reverse HTML entry point, or <see langword="null"/> when no reverse path exists.</summary>
    public string? ExportEntryPoint => TargetToHtml?.EntryPoint;
    /// <summary>Reverse conversion evidence contract, or <see langword="null"/> when unavailable.</summary>
    public string? ExportResultContract => TargetToHtml?.ResultContract;
    /// <summary>Compatibility alias for the HTML-to-target I/O boundary.</summary>
    [Obsolete("Use HtmlToTarget.IoAndAsyncBoundary or TargetToHtml.IoAndAsyncBoundary.")]
    public string IoAndAsyncBoundary => HtmlToTarget.IoAndAsyncBoundary;
    /// <summary>Compatibility aggregate of import and export profile names.</summary>
    [Obsolete("Use HtmlToTarget.Profiles or TargetToHtml.Profiles.")]
    public IReadOnlyList<string> Profiles => _profiles;
    /// <summary>Compatibility alias for HTML-to-target supported features.</summary>
    [Obsolete("Use HtmlToTarget.SupportedFeatures or TargetToHtml.SupportedFeatures.")]
    public IReadOnlyList<HtmlSemanticFeature> SupportedFeatures => HtmlToTarget.SupportedFeatures;
    /// <summary>Compatibility alias for HTML-to-target approximated features.</summary>
    [Obsolete("Use HtmlToTarget.ApproximatedFeatures or TargetToHtml.ApproximatedFeatures.")]
    public IReadOnlyList<HtmlSemanticFeature> ApproximatedFeatures => HtmlToTarget.ApproximatedFeatures;
    /// <summary>Compatibility alias for HTML-to-target unsupported features.</summary>
    [Obsolete("Use HtmlToTarget.UnsupportedFeatures or TargetToHtml.UnsupportedFeatures.")]
    public IReadOnlyList<HtmlSemanticFeature> UnsupportedFeatures => HtmlToTarget.UnsupportedFeatures;
    /// <summary>Whether the target exposes a reverse artifact-to-HTML route.</summary>
    public bool SupportsReverseHtml => TargetToHtml != null;

    /// <summary>Compatibility alias for the HTML-to-target feature outcome.</summary>
    [Obsolete("Use HtmlToTarget.GetSupport or TargetToHtml.GetSupport.")]
    public HtmlCapabilitySupportLevel GetSupport(HtmlSemanticFeature feature) {
        return HtmlToTarget.GetSupport(feature);
    }

    private static string Required(string value, string parameterName) {
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Value cannot be empty.", parameterName);
        return value.Trim();
    }

    private static string? Optional(string? value) => string.IsNullOrWhiteSpace(value) ? null : value!.Trim();

}
