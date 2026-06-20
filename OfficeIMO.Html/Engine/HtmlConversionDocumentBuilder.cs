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
        var styles = HtmlComputedStyleEngine.Compute(document);
        HtmlComputedStyleSummary styleSummary = HtmlComputedStyleEngine.Summarize(styles);
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest(document, options.ToResourcePipelineOptions());
        HtmlResourceDependencyPlan resourcePlan = HtmlResourceDependencyPlanner.Create(manifest);
        string normalized = options.IncludeNormalizedHtml
            ? HtmlNormalizer.Normalize(document, ConfigureNormalization(options))
            : string.Empty;

        return new HtmlConversionDocument(
            html,
            HtmlConversionProfileContracts.Get(options.Profile),
            logical,
            styles,
            styleSummary,
            manifest,
            resourcePlan,
            normalized);
    }

    private static HtmlNormalizationOptions ConfigureNormalization(HtmlConversionDocumentOptions options) {
        HtmlNormalizationOptions normalization = options.NormalizationOptions ?? new HtmlNormalizationOptions();
        if (normalization.BaseUri == null) {
            normalization.BaseUri = options.BaseUri;
        }

        normalization.UrlPolicy = options.UrlPolicy.Clone();
        normalization.UseBodyContentsOnly = options.UseBodyContentsOnly;
        return normalization;
    }
}
