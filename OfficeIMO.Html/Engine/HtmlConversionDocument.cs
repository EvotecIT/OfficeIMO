using AngleSharp.Dom;

namespace OfficeIMO.Html;

/// <summary>
/// Canonical OfficeIMO HTML conversion document shared by target adapters.
/// </summary>
public sealed class HtmlConversionDocument {
    internal HtmlConversionDocument(
        string sourceHtml,
        HtmlConversionProfileContract profileContract,
        HtmlLogicalDocument logicalDocument,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> computedStyles,
        HtmlComputedStyleSummary styleSummary,
        HtmlResourceManifest resourceManifest,
        HtmlResourceDependencyPlan resourcePlan,
        string normalizedHtml) {
        SourceHtml = sourceHtml ?? throw new ArgumentNullException(nameof(sourceHtml));
        ProfileContract = profileContract ?? throw new ArgumentNullException(nameof(profileContract));
        LogicalDocument = logicalDocument ?? throw new ArgumentNullException(nameof(logicalDocument));
        ComputedStyles = computedStyles ?? throw new ArgumentNullException(nameof(computedStyles));
        StyleSummary = styleSummary ?? throw new ArgumentNullException(nameof(styleSummary));
        ResourceManifest = resourceManifest ?? throw new ArgumentNullException(nameof(resourceManifest));
        ResourcePlan = resourcePlan ?? throw new ArgumentNullException(nameof(resourcePlan));
        NormalizedHtml = normalizedHtml ?? string.Empty;
    }

    /// <summary>Original HTML supplied by the caller.</summary>
    public string SourceHtml { get; }

    /// <summary>Profile contract advertised to target adapters and galleries.</summary>
    public HtmlConversionProfileContract ProfileContract { get; }

    /// <summary>Normalized logical structure used for semantic scoring and adapter planning.</summary>
    public HtmlLogicalDocument LogicalDocument { get; }

    /// <summary>Computed styles keyed by parsed source elements.</summary>
    public IReadOnlyDictionary<IElement, HtmlComputedStyle> ComputedStyles { get; }

    /// <summary>Compact computed-style capability summary.</summary>
    public HtmlComputedStyleSummary StyleSummary { get; }

    /// <summary>Raw resource manifest discovered in document order.</summary>
    public HtmlResourceManifest ResourceManifest { get; }

    /// <summary>Resource dependency plan grouped for adapters, reports, and gallery manifests.</summary>
    public HtmlResourceDependencyPlan ResourcePlan { get; }

    /// <summary>Policy-aware normalized HTML, or an empty string when normalization was disabled.</summary>
    public string NormalizedHtml { get; }

    /// <summary>HTML text target adapters should use when no adapter-specific source preference is configured.</summary>
    public string HtmlForConversion => SourceHtml;
}
