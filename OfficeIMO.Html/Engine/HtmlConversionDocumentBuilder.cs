using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

/// <summary>
/// Builds the shared OfficeIMO HTML conversion document consumed by target adapters.
/// </summary>
public static class HtmlConversionDocumentBuilder {
    /// <summary>
    /// Parses HTML and builds a shared conversion document with logical, style, resource, and normalized-output evidence.
    /// </summary>
    public static HtmlConversionDocument Build(string html, HtmlConversionDocumentOptions? options = null) {
        if (html == null) {
            throw new ArgumentNullException(nameof(html));
        }

        options ??= new HtmlConversionDocumentOptions();
        IHtmlDocument document = HtmlDocumentParser.ParseDocument(html);
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

        return new HtmlConversionDocument(
            html,
            HtmlConversionProfileContracts.Get(options.Profile),
            logical,
            styles,
            styleSummary,
            manifest,
            resourcePlan,
            normalized,
            adapterHtml);
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
        normalization.PreserveStyleElements = true;
        return normalization;
    }
}
