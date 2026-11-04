using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

public static partial class HtmlResourcePipeline {
    /// <summary>Plans generated backgrounds from the same effective contexts used by the painter.</summary>
    internal static HtmlResourceManifest BuildPageMarginBackgroundManifest(IHtmlDocument document,
        HtmlCssPageRuleSet pageRules, HtmlComputedStyleSet styles, HtmlResourcePipelineOptions options,
        CancellationToken cancellationToken, HtmlResourceSession? resources = null) {
        var manifest = new HtmlResourceManifest();
        IElement? root = document.DocumentElement;
        if (root == null || !styles.Elements.TryGetValue(root, out HtmlComputedStyle? rootStyle)) return manifest;
        Uri? baseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(document, options.BaseUri);
        var sources = new HashSet<string>(StringComparer.Ordinal);
        foreach (HtmlComputedStyle margin in pageRules.EnumerateMarginBackgroundStyles(rootStyle, styles, cancellationToken)) {
            foreach (string source in ExtractCssUrls(margin.GetValue("background-image"))) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!sources.Add(source)) continue;
                if (resources != null && resources.TryGetPlannedReference(source, out HtmlResourceReference planned)) manifest.Add(planned);
                else AddRaw(manifest, HtmlResourceKind.Image, root, "css-page-margin-background", source, baseUri, options);
            }
        }
        return manifest;
    }

    private static void AddPageMarginBackgroundResources(HtmlResourceManifest manifest, IHtmlDocument document,
        Uri? baseUri, HtmlResourcePipelineOptions options) {
        if (!document.QuerySelectorAll("style").Any(style => IsCssStyleElement(style)
            && IsApplicableMedia(style.GetAttribute("media") ?? string.Empty, options)
            && TryFindNextAtRule(StripCssCommentsOutsideStrings(style.TextContent), 0, "page", out _, out _))) return;
        var renderOptions = new HtmlRenderOptions {
            BaseUri = baseUri,
            DefaultFontSize = options.DefaultFontSize,
            CssMediaContextOverride = options.MediaContext,
            CssMediaWidthOverride = options.MediaWidth,
            CssMediaHeightOverride = options.MediaHeight,
            MediaFeatures = options.MediaFeatures.Clone()
        };
        if (options.MediaWidth.HasValue && options.MediaHeight.HasValue) {
            renderOptions.PageSize = new OfficePageSize(global::OfficeIMO.OfficeLength.Inches(options.MediaWidth.Value / HtmlRenderOptions.CssPixelsPerInch),
                global::OfficeIMO.OfficeLength.Inches(options.MediaHeight.Value / HtmlRenderOptions.CssPixelsPerInch));
        } else if (options.MediaContext == HtmlCssMediaContext.Screen) {
            renderOptions.PageSize = new OfficePageSize(global::OfficeIMO.OfficeLength.Inches(816D / HtmlRenderOptions.CssPixelsPerInch), global::OfficeIMO.OfficeLength.Inches(1056D / HtmlRenderOptions.CssPixelsPerInch));
        }
        HtmlCssPageRuleSet pageRules = HtmlCssPageSettingsResolver.Apply(document, renderOptions, new HtmlDiagnosticReport(), options.Limits);
        HtmlComputedStyleSet styles = HtmlComputedStyleEngine.ComputeForProvenance(document, options);
        foreach (HtmlResourceReference reference in BuildPageMarginBackgroundManifest(document, pageRules, styles, options, CancellationToken.None).Resources)
            manifest.Add(reference);
    }

    private static bool IsCssPageBackgroundReference(string css, int index) =>
        (GetCssDeclarationPropertyName(css, index) is "background" or "background-image") && IsInCssPageRule(css, index);

    private static bool IsInCssPageRule(string css, int index) =>
        IsInRanges(index, StructuralIndexes.GetValue(css, BuildStructuralIndex).GetPageRuleRanges(css));
}
