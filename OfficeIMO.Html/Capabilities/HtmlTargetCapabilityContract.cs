namespace OfficeIMO.Html;

/// <summary>Describes one built-in HTML target, its public APIs, and feature-level contract.</summary>
public sealed class HtmlTargetCapabilityContract {
    private readonly IReadOnlyList<HtmlSemanticFeature> _supported;
    private readonly IReadOnlyList<HtmlSemanticFeature> _approximated;
    private readonly IReadOnlyList<HtmlSemanticFeature> _unsupported;

    /// <summary>Creates a complete target capability contract.</summary>
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
        IEnumerable<HtmlSemanticFeature> unsupported) {
        Target = target;
        PackageName = Required(packageName, nameof(packageName));
        ArtifactName = Required(artifactName, nameof(artifactName));
        ImportEntryPoint = Required(importEntryPoint, nameof(importEntryPoint));
        ImportResultContract = Required(importResultContract, nameof(importResultContract));
        ExportEntryPoint = Optional(exportEntryPoint);
        ExportResultContract = Optional(exportResultContract);
        IoAndAsyncBoundary = Required(ioAndAsyncBoundary, nameof(ioAndAsyncBoundary));
        Profiles = ToStrings(profiles, nameof(profiles));
        _supported = ToFeatures(supported, nameof(supported));
        _approximated = ToFeatures(approximated, nameof(approximated));
        _unsupported = ToFeatures(unsupported, nameof(unsupported));
        ValidateCompleteFeaturePartition();
    }

    /// <summary>Target identifier used by preflight and adapter selection.</summary>
    public HtmlConversionTarget Target { get; }
    /// <summary>Package that owns the thin target projection.</summary>
    public string PackageName { get; }
    /// <summary>Native or rendered artifact produced by the target.</summary>
    public string ArtifactName { get; }
    /// <summary>Primary public HTML import entry point.</summary>
    public string ImportEntryPoint { get; }
    /// <summary>Public result or evidence contract returned by the import path.</summary>
    public string ImportResultContract { get; }
    /// <summary>Primary reverse HTML entry point, or <see langword="null"/> when no reverse path exists.</summary>
    public string? ExportEntryPoint { get; }
    /// <summary>Reverse conversion evidence contract, or <see langword="null"/> when unavailable.</summary>
    public string? ExportResultContract { get; }
    /// <summary>Documented path, stream, cancellation, and asynchronous boundary.</summary>
    public string IoAndAsyncBoundary { get; }
    /// <summary>Named import/export modes or profiles callers can select.</summary>
    public IReadOnlyList<string> Profiles { get; }
    /// <summary>Features represented through the documented target contract.</summary>
    public IReadOnlyList<HtmlSemanticFeature> SupportedFeatures => _supported;
    /// <summary>Features retained with a documented approximation.</summary>
    public IReadOnlyList<HtmlSemanticFeature> ApproximatedFeatures => _approximated;
    /// <summary>Features outside the current target contract.</summary>
    public IReadOnlyList<HtmlSemanticFeature> UnsupportedFeatures => _unsupported;
    /// <summary>Whether the target exposes a reverse artifact-to-HTML route.</summary>
    public bool SupportsReverseHtml => ExportEntryPoint != null;

    /// <summary>Gets the declared support outcome for one semantic feature.</summary>
    public HtmlCapabilitySupportLevel GetSupport(HtmlSemanticFeature feature) {
        if (_supported.Contains(feature)) return HtmlCapabilitySupportLevel.Supported;
        if (_approximated.Contains(feature)) return HtmlCapabilitySupportLevel.Approximated;
        if (_unsupported.Contains(feature)) return HtmlCapabilitySupportLevel.Unsupported;
        throw new ArgumentOutOfRangeException(nameof(feature), feature, "Unknown semantic feature.");
    }

    private void ValidateCompleteFeaturePartition() {
        var seen = new HashSet<HtmlSemanticFeature>();
        foreach (HtmlSemanticFeature feature in _supported.Concat(_approximated).Concat(_unsupported)) {
            if (!seen.Add(feature)) {
                throw new ArgumentException("Semantic feature '" + feature + "' was assigned more than once for " + Target + ".");
            }
        }

        HtmlSemanticFeature[] all = global::OfficeIMO.Internal.EnumCompat.GetValues<HtmlSemanticFeature>();
        if (seen.Count != all.Length) {
            string missing = string.Join(", ", all.Where(feature => !seen.Contains(feature)));
            throw new ArgumentException("Semantic features were not classified for " + Target + ": " + missing + ".");
        }
    }

    private static string Required(string value, string parameterName) {
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Value cannot be empty.", parameterName);
        return value.Trim();
    }

    private static string? Optional(string? value) => string.IsNullOrWhiteSpace(value) ? null : value!.Trim();

    private static IReadOnlyList<string> ToStrings(IEnumerable<string> values, string parameterName) {
        if (values == null) throw new ArgumentNullException(parameterName);
        return values.Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => value.Trim())
            .Distinct(StringComparer.Ordinal)
            .ToList()
            .AsReadOnly();
    }

    private static IReadOnlyList<HtmlSemanticFeature> ToFeatures(IEnumerable<HtmlSemanticFeature> values, string parameterName) {
        if (values == null) throw new ArgumentNullException(parameterName);
        return values.Distinct().OrderBy(value => value).ToList().AsReadOnly();
    }
}
