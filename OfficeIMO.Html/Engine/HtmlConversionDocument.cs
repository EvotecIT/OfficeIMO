using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

/// <summary>
/// Canonical OfficeIMO HTML conversion document shared by target adapters.
/// </summary>
public sealed partial class HtmlConversionDocument {
    private readonly IHtmlDocument _sourceDocumentForConversion;
    private readonly IHtmlDocument _defaultDocumentForConversion;

    internal HtmlConversionDocument(
        string sourceHtml,
        IHtmlDocument sourceDocumentForConversion,
        IHtmlDocument adapterDocument,
        IHtmlDocument documentForConversion,
        HtmlConversionProfileContract profileContract,
        HtmlInputTrust trust,
        HtmlLogicalDocument logicalDocument,
        HtmlComputedStyleSummary styleSummary,
        HtmlResourceManifest resourceManifest,
        HtmlResourceDependencyPlan resourcePlan,
        Uri? baseUri,
        Uri? fallbackBaseUri,
        string normalizedHtml,
        string adapterHtml) {
        SourceHtml = sourceHtml ?? throw new ArgumentNullException(nameof(sourceHtml));
        _sourceDocumentForConversion = sourceDocumentForConversion ?? throw new ArgumentNullException(nameof(sourceDocumentForConversion));
        AdapterDocument = adapterDocument ?? throw new ArgumentNullException(nameof(adapterDocument));
        _defaultDocumentForConversion = documentForConversion ?? throw new ArgumentNullException(nameof(documentForConversion));
        ProfileContract = profileContract ?? throw new ArgumentNullException(nameof(profileContract));
        Trust = trust;
        LogicalDocument = logicalDocument ?? throw new ArgumentNullException(nameof(logicalDocument));
        StyleSummary = styleSummary ?? throw new ArgumentNullException(nameof(styleSummary));
        ResourceManifest = resourceManifest ?? throw new ArgumentNullException(nameof(resourceManifest));
        ResourcePlan = resourcePlan ?? throw new ArgumentNullException(nameof(resourcePlan));
        BaseUri = baseUri;
        FallbackBaseUri = fallbackBaseUri;
        NormalizedHtml = normalizedHtml ?? string.Empty;
        AdapterHtml = adapterHtml ?? string.Empty;
    }

    /// <summary>Original HTML supplied by the caller.</summary>
    public string SourceHtml { get; }

    private IHtmlDocument AdapterDocument { get; }

    /// <summary>
    /// Creates an independent policy-normalized DOM for the conversion profile's default media context.
    /// </summary>
    public IHtmlDocument CreateDocumentForConversion() => HtmlDocumentParser.CloneDocument(_defaultDocumentForConversion);

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

    /// <summary>
    /// Creates an independent clone of the canonical source DOM for adapters that must apply
    /// their own element filters before URL resolution. Parsing remains owned by OfficeIMO.Html.
    /// </summary>
    internal IHtmlDocument CreateSourceDocumentForConversion() =>
        HtmlDocumentParser.CloneDocument(_sourceDocumentForConversion);

    /// <summary>
    /// Creates an unfiltered policy-normalized DOM for renderers that evaluate media queries
    /// against their concrete viewport or page dimensions.
    /// </summary>
    internal IHtmlDocument CreateDocumentForRendering() =>
        HtmlDocumentParser.CloneDocument(AdapterDocument);

    /// <summary>Profile contract advertised to target adapters and galleries.</summary>
    public HtmlConversionProfileContract ProfileContract { get; }

    /// <summary>Caller-assigned input trust boundary used by downstream adapters.</summary>
    public HtmlInputTrust Trust { get; }

    /// <summary>Normalized logical structure used for semantic scoring and adapter planning.</summary>
    public HtmlLogicalDocument LogicalDocument { get; }

    /// <summary>Compact computed-style capability summary.</summary>
    public HtmlComputedStyleSummary StyleSummary { get; }

    /// <summary>Raw resource manifest discovered in document order.</summary>
    public HtmlResourceManifest ResourceManifest { get; }

    /// <summary>Resource dependency plan grouped for adapters, reports, and gallery manifests.</summary>
    public HtmlResourceDependencyPlan ResourcePlan { get; }

    /// <summary>Effective base URI used to resolve relative resources, including a document <c>base</c> element when present.</summary>
    public Uri? BaseUri { get; }

    /// <summary>Caller-provided page URI before a document <c>base</c> element is applied.</summary>
    internal Uri? FallbackBaseUri { get; }

    /// <summary>Policy-aware normalized HTML, or an empty string when normalization was disabled.</summary>
    public string NormalizedHtml { get; }

    private string AdapterHtml { get; }

    /// <summary>HTML text target adapters should use when no adapter-specific source preference is configured.</summary>
    public string HtmlForConversion => string.IsNullOrWhiteSpace(AdapterHtml) ? SourceHtml : AdapterHtml;
}
