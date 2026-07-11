using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

/// <summary>
/// Canonical OfficeIMO HTML conversion document shared by target adapters.
/// </summary>
public sealed class HtmlConversionDocument {
    internal HtmlConversionDocument(
        string sourceHtml,
        IHtmlDocument sourceDocument,
        IHtmlDocument adapterDocument,
        IHtmlDocument documentForConversion,
        HtmlConversionProfileContract profileContract,
        HtmlInputTrust trust,
        HtmlLogicalDocument logicalDocument,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> computedStyles,
        HtmlComputedStyleSummary styleSummary,
        HtmlResourceManifest resourceManifest,
        HtmlResourceDependencyPlan resourcePlan,
        string normalizedHtml,
        string adapterHtml) {
        SourceHtml = sourceHtml ?? throw new ArgumentNullException(nameof(sourceHtml));
        SourceDocument = sourceDocument ?? throw new ArgumentNullException(nameof(sourceDocument));
        AdapterDocument = adapterDocument ?? throw new ArgumentNullException(nameof(adapterDocument));
        DocumentForConversion = documentForConversion ?? throw new ArgumentNullException(nameof(documentForConversion));
        ProfileContract = profileContract ?? throw new ArgumentNullException(nameof(profileContract));
        Trust = trust;
        LogicalDocument = logicalDocument ?? throw new ArgumentNullException(nameof(logicalDocument));
        ComputedStyles = computedStyles ?? throw new ArgumentNullException(nameof(computedStyles));
        StyleSummary = styleSummary ?? throw new ArgumentNullException(nameof(styleSummary));
        ResourceManifest = resourceManifest ?? throw new ArgumentNullException(nameof(resourceManifest));
        ResourcePlan = resourcePlan ?? throw new ArgumentNullException(nameof(resourcePlan));
        NormalizedHtml = normalizedHtml ?? string.Empty;
        AdapterHtml = adapterHtml ?? string.Empty;
    }

    /// <summary>Original HTML supplied by the caller.</summary>
    public string SourceHtml { get; }

    /// <summary>Parsed source DOM used by logical, style, and resource analysis.</summary>
    public IHtmlDocument SourceDocument { get; }

    /// <summary>Policy-normalized DOM filtered for the conversion profile's default media context.</summary>
    public IHtmlDocument DocumentForConversion { get; }

    private IHtmlDocument AdapterDocument { get; }

    /// <summary>
    /// Creates a policy-normalized DOM filtered for a target media context without reparsing source HTML or mutating shared state.
    /// </summary>
    /// <param name="mediaContext">Screen or print media context selected by the target adapter.</param>
    /// <returns>An independent DOM clone that the target adapter may safely mutate.</returns>
    public IHtmlDocument CreateDocumentForConversion(HtmlCssMediaContext mediaContext) {
        IHtmlDocument document = HtmlDocumentParser.CloneDocument(AdapterDocument);
        HtmlActiveMediaFilter.Filter(document, mediaContext);
        return document;
    }

    /// <summary>Profile contract advertised to target adapters and galleries.</summary>
    public HtmlConversionProfileContract ProfileContract { get; }

    /// <summary>Caller-assigned input trust boundary used by downstream adapters.</summary>
    public HtmlInputTrust Trust { get; }

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

    private string AdapterHtml { get; }

    /// <summary>HTML text target adapters should use when no adapter-specific source preference is configured.</summary>
    public string HtmlForConversion => string.IsNullOrWhiteSpace(AdapterHtml) ? SourceHtml : AdapterHtml;
}
