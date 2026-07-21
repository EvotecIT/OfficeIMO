using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

/// <summary>
/// Canonical OfficeIMO HTML conversion document shared by target adapters.
/// </summary>
public sealed partial class HtmlConversionDocument {
    private readonly IHtmlDocument _sourceDocument;
    private readonly HtmlConversionDocumentOptions _options;
    private readonly HtmlCssMediaContext _mediaContext;
    private readonly Lazy<IHtmlDocument> _adapterDocument;
    private readonly Lazy<HtmlLogicalDocument> _logicalDocument;
    private readonly Lazy<HtmlSemanticDocument> _semanticDocument;
    private readonly Lazy<HtmlComputedStyleSummary> _styleSummary;
    private readonly Lazy<HtmlResourceManifest> _resourceManifest;
    private readonly Lazy<HtmlResourceDependencyPlan> _resourcePlan;
    private readonly Lazy<string> _normalizedHtml;
    private readonly HtmlDiagnosticReport _diagnostics = new HtmlDiagnosticReport();
    private readonly object _diagnosticSync = new object();
    private readonly object _analysisSync = new object();
    private readonly Dictionary<HtmlConversionTarget, HtmlConversionPreflight> _preflight = new Dictionary<HtmlConversionTarget, HtmlConversionPreflight>();

    internal HtmlConversionDocument(
        string sourceHtml,
        IHtmlDocument sourceDocument,
        HtmlConversionDocumentOptions options,
        Uri? baseUri) {
        SourceHtml = sourceHtml ?? throw new ArgumentNullException(nameof(sourceHtml));
        _sourceDocument = sourceDocument ?? throw new ArgumentNullException(nameof(sourceDocument));
        _options = options?.Clone() ?? throw new ArgumentNullException(nameof(options));
        _mediaContext = _options.Profile == HtmlConversionProfile.HighFidelityPrint
            ? HtmlCssMediaContext.Print
            : HtmlCssMediaContext.Screen;
        ProfileContract = HtmlConversionProfileContracts.Get(_options.Profile);
        Trust = _options.Trust;
        BaseUri = baseUri;
        FallbackBaseUri = _options.BaseUri;
        _adapterDocument = new Lazy<IHtmlDocument>(BuildAdapterDocument, LazyThreadSafetyMode.ExecutionAndPublication);
        _logicalDocument = new Lazy<HtmlLogicalDocument>(
            () => AnalyzeSource(() => HtmlLogicalDocumentBuilder.FromDocument(_sourceDocument, _options.UseBodyContentsOnly)),
            LazyThreadSafetyMode.ExecutionAndPublication);
        _semanticDocument = new Lazy<HtmlSemanticDocument>(
            () => AnalyzeSource(() => HtmlSemanticDocumentBuilder.FromDocument(
                _adapterDocument.Value,
                _mediaContext,
                _options.Limits)),
            LazyThreadSafetyMode.ExecutionAndPublication);
        _styleSummary = new Lazy<HtmlComputedStyleSummary>(
            () => AnalyzeSource(() => HtmlComputedStyleEngine.Summarize(HtmlComputedStyleEngine.Compute(_sourceDocument, _mediaContext, _options.Limits))),
            LazyThreadSafetyMode.ExecutionAndPublication);
        _resourceManifest = new Lazy<HtmlResourceManifest>(
            () => AnalyzeSource(() => HtmlResourcePipeline.BuildManifest(_sourceDocument, _options.ToResourcePipelineOptions())),
            LazyThreadSafetyMode.ExecutionAndPublication);
        _resourcePlan = new Lazy<HtmlResourceDependencyPlan>(
            () => HtmlResourceDependencyPlanner.Create(_resourceManifest.Value),
            LazyThreadSafetyMode.ExecutionAndPublication);
        _normalizedHtml = new Lazy<string>(
            () => _options.IncludeNormalizedHtml
                ? AnalyzeSource(() => HtmlNormalizer.Normalize(_sourceDocument, ConfigureNormalization(_sourceDocument, _options)))
                : string.Empty,
            LazyThreadSafetyMode.ExecutionAndPublication);
    }

    /// <summary>Original HTML supplied by the caller.</summary>
    public string SourceHtml { get; }

    /// <summary>
    /// Creates an independent policy-normalized DOM for the conversion profile's default media context.
    /// </summary>
    public IHtmlDocument CreateDocumentForConversion() => CreateDocumentForConversion(_mediaContext);

    /// <summary>
    /// Creates a policy-normalized DOM filtered for a target media context without reparsing source HTML or mutating shared state.
    /// </summary>
    /// <param name="mediaContext">Screen or print media context selected by the target adapter.</param>
    /// <returns>An independent DOM clone that the target adapter may safely mutate.</returns>
    public IHtmlDocument CreateDocumentForConversion(HtmlCssMediaContext mediaContext) {
        IHtmlDocument canonical = _adapterDocument.Value;
        IHtmlDocument document;
        lock (_analysisSync) document = HtmlDocumentParser.CloneDocument(canonical);
        var diagnostics = new HtmlDiagnosticReport();
        HtmlActiveMediaFilter.Filter(document, mediaContext, diagnostics);
        if (diagnostics.Count > 0) {
            lock (_diagnosticSync) _diagnostics.AddRange(diagnostics);
        }
        return document;
    }

    /// <summary>
    /// Builds the shared semantic projection in the media context requested by a native target adapter.
    /// </summary>
    internal HtmlSemanticDocument CreateSemanticDocumentForConversion(HtmlCssMediaContext mediaContext) {
        if (!Enum.IsDefined(typeof(HtmlCssMediaContext), mediaContext)) throw new ArgumentOutOfRangeException(nameof(mediaContext));
        if (mediaContext == _mediaContext) return _semanticDocument.Value;
        IHtmlDocument document = CreateDocumentForConversion(mediaContext);
        return AnalyzeSource(() => HtmlSemanticDocumentBuilder.FromDocument(document, mediaContext, _options.Limits));
    }

    /// <summary>
    /// Creates an independent clone of the canonical source DOM for adapters that must apply
    /// their own element filters before URL resolution. Parsing remains owned by OfficeIMO.Html.
    /// </summary>
    internal IHtmlDocument CreateSourceDocumentForConversion() {
        lock (_analysisSync) return HtmlDocumentParser.CloneDocument(_sourceDocument);
    }

    /// <summary>
    /// Creates a policy-normalized adapter DOM without selecting a media context. Structural
    /// adapters can preserve responsive alternatives while still inheriting the document's trust
    /// and URL decisions.
    /// </summary>
    internal IHtmlDocument CreatePolicyNormalizedDocumentForConversion() {
        IHtmlDocument canonical = _adapterDocument.Value;
        lock (_analysisSync) return HtmlDocumentParser.CloneDocument(canonical);
    }

    /// <summary>
    /// Creates an unfiltered source DOM for renderers that apply their own concrete viewport,
    /// hyperlink, and resource policies. This keeps an explicitly configured resolver capable of
    /// authorizing a source that a generic untrusted adapter would omit.
    /// </summary>
    internal IHtmlDocument CreateDocumentForRendering() {
        lock (_analysisSync) return HtmlDocumentParser.CloneDocument(_sourceDocument);
    }

    /// <summary>Shared limits snapshot used by target adapters and renderers.</summary>
    internal HtmlConversionLimits Limits => _options.Limits.Clone();

    /// <summary>Default media context selected by the shared conversion profile.</summary>
    internal HtmlCssMediaContext MediaContext => _mediaContext;

    /// <summary>Hyperlink policy snapshot supplied at the shared input boundary.</summary>
    internal HtmlUrlPolicy HyperlinkUrlPolicy => _options.UrlPolicy.Clone();

    /// <summary>Resource policy snapshot supplied at the shared input boundary.</summary>
    internal HtmlUrlPolicy ResourceUrlPolicy => _options.ResourceUrlPolicy.Clone();

    /// <summary>Profile contract advertised to target adapters and galleries.</summary>
    public HtmlConversionProfileContract ProfileContract { get; }

    /// <summary>Caller-assigned input trust boundary used by downstream adapters.</summary>
    public HtmlInputTrust Trust { get; }

    /// <summary>Normalized logical structure used for semantic scoring and adapter planning.</summary>
    public HtmlLogicalDocument LogicalDocument => _logicalDocument.Value;

    /// <summary>Typed semantic structure interpreted once for generic native target adapters.</summary>
    public HtmlSemanticDocument SemanticDocument => _semanticDocument.Value;

    /// <summary>
    /// Predicts supported, approximated, and omitted source features for a target before artifact creation.
    /// Results are cached per prepared document and target.
    /// </summary>
    public HtmlConversionPreflight AnalyzeFor(HtmlConversionTarget target) {
        if (!Enum.IsDefined(typeof(HtmlConversionTarget), target)) throw new ArgumentOutOfRangeException(nameof(target));
        lock (_analysisSync) {
            if (_preflight.TryGetValue(target, out HtmlConversionPreflight? result)) return result;
            result = HtmlConversionPreflightAnalyzer.Analyze(this, target);
            _preflight[target] = result;
            return result;
        }
    }

    /// <summary>Compact computed-style capability summary.</summary>
    public HtmlComputedStyleSummary StyleSummary => _styleSummary.Value;

    /// <summary>Raw resource manifest discovered in document order.</summary>
    public HtmlResourceManifest ResourceManifest => _resourceManifest.Value;

    /// <summary>Resource dependency plan grouped for adapters, reports, and gallery manifests.</summary>
    public HtmlResourceDependencyPlan ResourcePlan => _resourcePlan.Value;

    /// <summary>Effective base URI used to resolve relative resources, including a document <c>base</c> element when present.</summary>
    public Uri? BaseUri { get; }

    /// <summary>Caller-provided page URI before a document <c>base</c> element is applied.</summary>
    internal Uri? FallbackBaseUri { get; }

    /// <summary>Policy-aware normalized HTML, or an empty string when normalization was disabled.</summary>
    public string NormalizedHtml => _normalizedHtml.Value;

    /// <summary>Diagnostics emitted while lazily preparing shared conversion views.</summary>
    public IReadOnlyList<HtmlDiagnostic> Diagnostics {
        get {
            lock (_diagnosticSync) return _diagnostics.Diagnostics.ToArray();
        }
    }

    /// <summary>HTML text target adapters should use when no adapter-specific source preference is configured.</summary>
    public string HtmlForConversion {
        get {
            IHtmlDocument document = CreateDocumentForConversion();
            return document.DocumentElement?.OuterHtml ?? SourceHtml;
        }
    }

    private IHtmlDocument BuildAdapterDocument() {
        return AnalyzeSource(() => {
            string adapterHtml = HtmlNormalizer.Normalize(_sourceDocument, ConfigureAdapterNormalization(_sourceDocument, _options));
            IHtmlDocument document = HtmlDocumentParser.ParseDocument(adapterHtml);
            HtmlConversionInputGuard.ValidateDocument(document, _options.Limits);
            return document;
        });
    }

    private T AnalyzeSource<T>(Func<T> analysis) {
        lock (_analysisSync) return analysis();
    }
}
