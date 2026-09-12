using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

public sealed partial class HtmlConversionDocument {
    /// <summary>
    /// Parses HTML and builds a shared conversion document with logical, style, resource, and normalized-output evidence.
    /// </summary>
    public static HtmlConversionDocument Parse(string html, HtmlConversionDocumentOptions? options = null) => Parse(html, options, CancellationToken.None);

    /// <summary>Parses inert HTML with explicit cooperative cancellation.</summary>
    public static HtmlConversionDocument Parse(string html, HtmlConversionDocumentOptions? options, CancellationToken cancellationToken) {
        if (html == null) {
            throw new ArgumentNullException(nameof(html));
        }

        HtmlConversionDocumentOptions resolved = options?.Clone() ?? new HtmlConversionDocumentOptions();
        resolved.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        HtmlConversionInputGuard.ValidateSource(html, resolved.Limits);
        // Keep the existing provider's retained native tree until owned nodes are requested.
        // This avoids eagerly duplicating the DOM for render-only and semantic-only callers.
        Dom.HtmlDocument? sourceSnapshot = null;
        IHtmlDocument document;
        if (ReferenceEquals(resolved.ParserProvider, Providers.AngleSharpHtmlParser.Instance)) {
            document = HtmlDocumentParser.ParseDocument(html, cancellationToken);
        } else {
            try {
                sourceSnapshot = resolved.ParserProvider.Parse(html, new Dom.HtmlParseOptions {
                    MaxInputCharacters = resolved.Limits.MaxInputCharacters,
                    MaxNodes = resolved.Limits.MaxHtmlNodes,
                    MaxDepth = resolved.Limits.MaxHtmlDepth
                }, cancellationToken) ?? throw new InvalidOperationException("The HTML parser provider returned no document.");
            } catch (Dom.HtmlParseLimitException exception) when (
                exception.LimitName == nameof(Dom.HtmlParseOptions.MaxInputCharacters)
                || exception.LimitName == nameof(Dom.HtmlParseOptions.MaxNodes)
                || exception.LimitName == nameof(Dom.HtmlParseOptions.MaxDepth)) {
                string source = exception.LimitName == nameof(Dom.HtmlParseOptions.MaxNodes) ? nameof(HtmlConversionLimits.MaxHtmlNodes)
                    : exception.LimitName == nameof(Dom.HtmlParseOptions.MaxDepth) ? nameof(HtmlConversionLimits.MaxHtmlDepth)
                    : nameof(HtmlConversionLimits.MaxInputCharacters);
                string code = exception.LimitName == nameof(Dom.HtmlParseOptions.MaxNodes) ? HtmlRenderDiagnosticCodes.NodeLimitExceeded
                    : exception.LimitName == nameof(Dom.HtmlParseOptions.MaxDepth) ? HtmlConversionDiagnosticCodes.HtmlDepthLimitExceeded
                    : HtmlRenderDiagnosticCodes.InputCharacterLimitExceeded;
                throw new HtmlDomLimitException(code, "The HTML parser exceeded a shared conversion limit.", source, exception.Actual, exception.Maximum, exception);
            }
            sourceSnapshot = HtmlConversionInputGuard.CaptureOwnedTree(sourceSnapshot, resolved.Limits, cancellationToken);
            document = NativeDomBridge.GetNativeDocument(sourceSnapshot, cancellationToken);
        }
        cancellationToken.ThrowIfCancellationRequested();
        HtmlConversionInputGuard.ValidateDocument(document, resolved.Limits, cancellationToken);
        Uri? baseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(document, resolved.BaseUri);
        cancellationToken.ThrowIfCancellationRequested();
        return new HtmlConversionDocument(html, document, resolved, baseUri, sourceSnapshot);
    }

    /// <summary>Creates a conversion snapshot from an owned tree without reparsing its serialized markup.</summary>
    public static HtmlConversionDocument FromDocument(Dom.HtmlDocument document, HtmlConversionDocumentOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        HtmlConversionDocumentOptions resolved = options?.Clone() ?? new HtmlConversionDocumentOptions();
        resolved.Validate();
        Dom.HtmlDocument snapshot = HtmlConversionInputGuard.CaptureOwnedTree(document, resolved.Limits, CancellationToken.None);
        IHtmlDocument native = NativeDomBridge.GetNativeDocument(snapshot);
        HtmlConversionInputGuard.ValidateDocument(native, resolved.Limits);
        string source = HtmlConversionSourceWriter.Serialize(native, resolved.Limits);
        return new HtmlConversionDocument(source, native, resolved, HtmlDocumentParser.ResolveEffectiveBaseUri(native, resolved.BaseUri), snapshot);
    }

    /// <summary>Edits an independent source snapshot and retains this conversion's trust, resource and fidelity options.</summary>
    public HtmlConversionDocument Edit(Action<Dom.HtmlDocument> edit) {
        if (edit == null) throw new ArgumentNullException(nameof(edit));
        Dom.HtmlDocument clone = Document.CloneAttached();
        try { edit(clone); return FromDocument(clone, _options); }
        finally { clone.Freeze(); }
    }

    /// <summary>
    /// Gives low-level analysis helpers an independent source DOM while retaining the same bounded
    /// parser entry point as converters. The returned DOM may be safely mutated by the caller.
    /// </summary>
    internal static IHtmlDocument ParseSourceDocumentForAnalysis(string html) =>
        ParseSourceDocumentForAnalysis(html, HtmlConversionLimits.CreateUntrustedProfile());

    /// <summary>Parses a low-level analysis clone using the caller's resolved shared limits.</summary>
    internal static IHtmlDocument ParseSourceDocumentForAnalysis(string html, HtmlConversionLimits limits) =>
        Parse(
            html,
            new HtmlConversionDocumentOptions {
                IncludeNormalizedHtml = false,
                Limits = (limits ?? HtmlConversionLimits.CreateUntrustedProfile()).Clone()
            })
        .CreateSourceDocumentForConversion();

    private static HtmlNormalizationOptions ConfigureNormalization(IHtmlDocument document, HtmlConversionDocumentOptions options) {
        HtmlNormalizationOptions source = options.NormalizationOptions ?? new HtmlNormalizationOptions();
        return new HtmlNormalizationOptions {
            BaseUri = source.BaseUri ?? HtmlDocumentParser.ResolveEffectiveBaseUri(document, options.BaseUri),
            BaseElementBaseUri = source.BaseElementBaseUri ?? source.BaseUri ?? options.BaseUri,
            UrlPolicy = (options.UrlPolicy ?? HtmlUrlPolicy.CreateOfficeIMOProfile()).Clone(),
            ResourceUrlPolicy = (options.ResourceUrlPolicy ?? HtmlUrlPolicy.CreateEmbeddedResourceProfile()).Clone(),
            Limits = options.Limits.Clone(),
            MaxResponsiveImageCandidates = options.Limits.MaxResponsiveImageCandidates,
            UseBodyContentsOnly = options.UseBodyContentsOnly,
            PreserveComments = source.PreserveComments,
            PreserveStyleElements = source.PreserveStyleElements,
            RemoveEventHandlerAttributes = source.RemoveEventHandlerAttributes,
            CollapseTextWhitespace = source.CollapseTextWhitespace
        };
    }

    private static HtmlNormalizationOptions ConfigureAdapterNormalization(IHtmlDocument document, HtmlConversionDocumentOptions options) {
        HtmlNormalizationOptions normalization = ConfigureNormalization(document, options);
        HtmlNormalizationOptions source = options.NormalizationOptions ?? new HtmlNormalizationOptions();
        normalization.BaseElementBaseUri = source.BaseElementBaseUri ?? source.BaseUri ?? options.BaseUri;
        normalization.UseBodyContentsOnly = false;
        // Target adapters still need source comments and significant whitespace so they can
        // apply their own supported-feature and diagnostic policies. The normalized review
        // representation may remain compact, but the adapter DOM must not be lossy.
        normalization.PreserveComments = true;
        normalization.PreserveSkippedElementMarkers = true;
        normalization.PreserveStyleElements = true;
        normalization.CollapseTextWhitespace = false;
        return normalization;
    }
}
