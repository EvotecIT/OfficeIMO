namespace OfficeIMO.Html;

/// <summary>Describes one directional HTML conversion route and its feature contract.</summary>
public abstract class HtmlConversionRouteCapabilityContract {
    private readonly IReadOnlyList<HtmlSemanticFeature> _supported;
    private readonly IReadOnlyList<HtmlSemanticFeature> _approximated;
    private readonly IReadOnlyList<HtmlSemanticFeature> _unsupported;

    /// <summary>Creates a directional route capability contract.</summary>
    protected HtmlConversionRouteCapabilityContract(
        string entryPoint,
        string resultContract,
        string ioAndAsyncBoundary,
        string diagnosticsContract,
        IEnumerable<string> profiles,
        IEnumerable<HtmlSemanticFeature> supported,
        IEnumerable<HtmlSemanticFeature> approximated,
        IEnumerable<HtmlSemanticFeature> unsupported) {
        EntryPoint = Required(entryPoint, nameof(entryPoint));
        ResultContract = Required(resultContract, nameof(resultContract));
        IoAndAsyncBoundary = Required(ioAndAsyncBoundary, nameof(ioAndAsyncBoundary));
        DiagnosticsContract = Required(diagnosticsContract, nameof(diagnosticsContract));
        Profiles = ToStrings(profiles, nameof(profiles));
        _supported = ToFeatures(supported, nameof(supported));
        _approximated = ToFeatures(approximated, nameof(approximated));
        _unsupported = ToFeatures(unsupported, nameof(unsupported));
        ValidateCompleteFeaturePartition();
    }

    /// <summary>Primary public conversion entry point.</summary>
    public string EntryPoint { get; }

    /// <summary>Public result or evidence contract returned by the route.</summary>
    public string ResultContract { get; }

    /// <summary>Documented path, stream, cancellation, and asynchronous boundary.</summary>
    public string IoAndAsyncBoundary { get; }

    /// <summary>Structured diagnostics and per-construct fidelity evidence exposed by this direction.</summary>
    public string DiagnosticsContract { get; }

    /// <summary>Named modes or profiles callers can select for this direction.</summary>
    public IReadOnlyList<string> Profiles { get; }

    /// <summary>Features represented through the route's documented contract.</summary>
    public IReadOnlyList<HtmlSemanticFeature> SupportedFeatures => _supported;

    /// <summary>Features retained with a documented approximation.</summary>
    public IReadOnlyList<HtmlSemanticFeature> ApproximatedFeatures => _approximated;

    /// <summary>Features outside the route's current contract.</summary>
    public IReadOnlyList<HtmlSemanticFeature> UnsupportedFeatures => _unsupported;

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
                throw new ArgumentException("Semantic feature '" + feature + "' was assigned more than once.");
            }
        }

        HtmlSemanticFeature[] all = global::OfficeIMO.Internal.EnumCompat.GetValues<HtmlSemanticFeature>();
        if (seen.Count != all.Length) {
            string missing = string.Join(", ", all.Where(feature => !seen.Contains(feature)));
            throw new ArgumentException("Semantic features were not classified: " + missing + ".");
        }
    }

    private static string Required(string value, string parameterName) {
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Value cannot be empty.", parameterName);
        return value.Trim();
    }

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

/// <summary>Capability contract for an HTML-to-native-target route.</summary>
public sealed class HtmlToTargetCapabilityContract : HtmlConversionRouteCapabilityContract {
    /// <summary>Creates an HTML-to-target route contract.</summary>
    public HtmlToTargetCapabilityContract(
        string entryPoint,
        string resultContract,
        string ioAndAsyncBoundary,
        string diagnosticsContract,
        IEnumerable<string> profiles,
        IEnumerable<HtmlSemanticFeature> supported,
        IEnumerable<HtmlSemanticFeature> approximated,
        IEnumerable<HtmlSemanticFeature> unsupported)
        : base(entryPoint, resultContract, ioAndAsyncBoundary, diagnosticsContract, profiles,
            supported, approximated, unsupported) { }
}

/// <summary>Capability contract for a native-target-to-HTML route.</summary>
public sealed class TargetToHtmlCapabilityContract : HtmlConversionRouteCapabilityContract {
    /// <summary>Creates a target-to-HTML route contract.</summary>
    public TargetToHtmlCapabilityContract(
        string entryPoint,
        string resultContract,
        string ioAndAsyncBoundary,
        string diagnosticsContract,
        IEnumerable<string> profiles,
        IEnumerable<HtmlSemanticFeature> supported,
        IEnumerable<HtmlSemanticFeature> approximated,
        IEnumerable<HtmlSemanticFeature> unsupported)
        : base(entryPoint, resultContract, ioAndAsyncBoundary, diagnosticsContract, profiles,
            supported, approximated, unsupported) { }
}
