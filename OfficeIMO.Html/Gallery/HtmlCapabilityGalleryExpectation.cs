namespace OfficeIMO.Html;

/// <summary>
/// Expected proof outcome for one capability or known simplification in an HTML conversion gallery scenario.
/// </summary>
public enum HtmlCapabilityGalleryExpectationOutcome {
    /// <summary>The feature is expected to survive the conversion path.</summary>
    Preserved,

    /// <summary>The feature is expected to be intentionally simplified by the selected profile.</summary>
    Simplified,

    /// <summary>The feature is expected to be blocked by policy or safety limits.</summary>
    Blocked,

    /// <summary>The feature is expected to be omitted because it is outside the selected profile contract.</summary>
    Omitted,

    /// <summary>The feature is expected to be reported through diagnostics rather than silently accepted.</summary>
    Reported,

    /// <summary>The scenario is expected to provide text/logical readback evidence.</summary>
    TextProof,

    /// <summary>The scenario is expected to provide rendered, visual, raster, SVG, or positioned review evidence.</summary>
    VisualProof
}

/// <summary>
/// Describes one expected proof item in an HTML capability-gallery manifest.
/// </summary>
public sealed class HtmlCapabilityGalleryExpectation {
    /// <summary>
    /// Creates a proof expectation.
    /// </summary>
    /// <param name="feature">Stable feature or behavior name.</param>
    /// <param name="outcome">Expected proof outcome.</param>
    /// <param name="evidence">Expected artifact, diagnostic, score, or readback signal proving the outcome.</param>
    public HtmlCapabilityGalleryExpectation(string feature, HtmlCapabilityGalleryExpectationOutcome outcome, string evidence) {
        Feature = feature ?? throw new ArgumentNullException(nameof(feature));
        Outcome = outcome;
        Evidence = evidence ?? throw new ArgumentNullException(nameof(evidence));
    }

    /// <summary>Stable feature or behavior name.</summary>
    public string Feature { get; }

    /// <summary>Expected proof outcome.</summary>
    public HtmlCapabilityGalleryExpectationOutcome Outcome { get; }

    /// <summary>Expected artifact, diagnostic, score, or readback signal proving the outcome.</summary>
    public string Evidence { get; }
}
