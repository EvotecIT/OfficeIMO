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

        options ??= new HtmlConversionDocumentOptions();
        IHtmlDocument document = HtmlDocumentParser.ParseDocument(html);
        Uri? baseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(document, options.BaseUri);
        HtmlLogicalDocument logical = HtmlLogicalDocumentBuilder.FromDocument(document, options.UseBodyContentsOnly);
        HtmlCssMediaContext mediaContext = options.Profile == HtmlConversionProfile.HighFidelityPrint
            ? HtmlCssMediaContext.Print
            : HtmlCssMediaContext.Screen;
        var styles = HtmlComputedStyleEngine.Compute(document, mediaContext);
        HtmlComputedStyleSummary styleSummary = HtmlComputedStyleEngine.Summarize(styles);
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest(document, options.ToResourcePipelineOptions());
        HtmlResourceDependencyPlan resourcePlan = HtmlResourceDependencyPlanner.Create(manifest);
        string normalized = options.IncludeNormalizedHtml
            ? HtmlNormalizer.Normalize(document, ConfigureNormalization(document, options))
            : string.Empty;
        string adapterHtml = HtmlNormalizer.Normalize(document, ConfigureAdapterNormalization(document, options));
        IHtmlDocument adapterDocument = HtmlDocumentParser.ParseDocument(adapterHtml);
        IHtmlDocument documentForConversion = HtmlDocumentParser.CloneDocument(adapterDocument);
        HtmlActiveMediaFilter.Filter(documentForConversion, mediaContext);
        string filteredAdapterHtml = documentForConversion.DocumentElement?.OuterHtml ?? adapterHtml;

        return new HtmlConversionDocument(
            html,
            HtmlDocumentParser.CloneDocument(document),
            adapterDocument,
            documentForConversion,
            HtmlConversionProfileContracts.Get(options.Profile),
            options.Trust,
            logical,
            styleSummary,
            manifest,
            resourcePlan,
            baseUri,
            options.BaseUri,
            normalized,
            filteredAdapterHtml);
    }

    private static HtmlNormalizationOptions ConfigureNormalization(IHtmlDocument document, HtmlConversionDocumentOptions options) {
        HtmlNormalizationOptions source = options.NormalizationOptions ?? new HtmlNormalizationOptions();
        return new HtmlNormalizationOptions {
            BaseUri = source.BaseUri ?? HtmlDocumentParser.ResolveEffectiveBaseUri(document, options.BaseUri),
            BaseElementBaseUri = source.BaseElementBaseUri ?? source.BaseUri ?? options.BaseUri,
            UrlPolicy = (options.UrlPolicy ?? HtmlUrlPolicy.CreateOfficeIMOProfile()).Clone(),
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
