using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

public sealed partial class HtmlConversionDocument {
    /// <summary>
    /// Parses HTML and builds a shared conversion document with logical, style, resource, and normalized-output evidence.
    /// </summary>
    public static HtmlConversionDocument Parse(string html, HtmlConversionDocumentOptions? options = null) {
        if (html == null) {
            throw new ArgumentNullException(nameof(html));
        }

        HtmlConversionDocumentOptions resolved = (options ?? new HtmlConversionDocumentOptions()).Clone();
        resolved.Validate();
        HtmlConversionInputGuard.ValidateSource(html, resolved.Limits);
        IHtmlDocument document = HtmlDocumentParser.ParseDocument(html);
        HtmlConversionInputGuard.ValidateDocument(document, resolved.Limits);
        Uri? baseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(document, resolved.BaseUri);
        return new HtmlConversionDocument(html, document, resolved, baseUri);
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
