namespace OfficeIMO.Html;

/// <summary>Predicted handling of a source feature before target artifact creation.</summary>
public enum HtmlConversionPreflightOutcome {
    /// <summary>The target contract represents the feature.</summary>
    Supported,
    /// <summary>The target contract retains the feature through a documented approximation.</summary>
    Approximated,
    /// <summary>The target contract omits the feature.</summary>
    Omitted
}

/// <summary>Preflight prediction for one semantic feature.</summary>
public sealed class HtmlFeaturePreflightResult {
    internal HtmlFeaturePreflightResult(HtmlSemanticFeature feature, bool isPresent, int occurrenceCount,
        HtmlConversionPreflightOutcome outcome, HtmlSemanticSourceLocation? firstSourceLocation) {
        Feature = feature;
        IsPresent = isPresent;
        OccurrenceCount = occurrenceCount;
        Outcome = outcome;
        FirstSourceLocation = firstSourceLocation;
    }

    /// <summary>Semantic feature.</summary>
    public HtmlSemanticFeature Feature { get; }
    /// <summary>Whether the source contains evidence of the feature.</summary>
    public bool IsPresent { get; }
    /// <summary>Best-effort semantic occurrence count.</summary>
    public int OccurrenceCount { get; }
    /// <summary>Predicted target outcome.</summary>
    public HtmlConversionPreflightOutcome Outcome { get; }
    /// <summary>First source location associated with the feature.</summary>
    public HtmlSemanticSourceLocation? FirstSourceLocation { get; }
}

/// <summary>Executable, source-aware target capability analysis performed before artifact creation.</summary>
public sealed class HtmlConversionPreflight {
    internal HtmlConversionPreflight(HtmlConversionTarget target, HtmlTargetCapabilityContract contract,
        IReadOnlyList<HtmlFeaturePreflightResult> features, IReadOnlyList<HtmlDiagnostic> diagnostics) {
        Target = target;
        Contract = contract;
        Features = features;
        Diagnostics = diagnostics;
    }

    /// <summary>Analyzed target.</summary>
    public HtmlConversionTarget Target { get; }
    /// <summary>Capability contract used for prediction.</summary>
    public HtmlTargetCapabilityContract Contract { get; }
    /// <summary>Complete feature predictions in enum order.</summary>
    public IReadOnlyList<HtmlFeaturePreflightResult> Features { get; }
    /// <summary>Approximation and omission diagnostics with source-to-target provenance.</summary>
    public IReadOnlyList<HtmlDiagnostic> Diagnostics { get; }
    /// <summary>Whether any present feature is predicted to be approximated or omitted.</summary>
    public bool HasPotentialLoss => Features.Any(feature => feature.IsPresent && feature.Outcome != HtmlConversionPreflightOutcome.Supported);

    /// <summary>Gets one feature prediction.</summary>
    public HtmlFeaturePreflightResult Get(HtmlSemanticFeature feature) =>
        Features.First(item => item.Feature == feature);
}
