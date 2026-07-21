using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

/// <summary>
/// Shared resource discovery and policy planning for OfficeIMO HTML workflows.
/// </summary>
public static partial class HtmlResourcePipeline {
    private const int MaxCustomPropertyResolutionDepth = 8;
    private const string ResourceSelector = "image, meta[http-equiv], [src], [srcset], [href], [xlink\\:href], [data], [data-src], [data-original], [data-original-src], [data-lazy-src], [data-srcset], [data-original-srcset], [data-lazy-srcset], [poster], [data-poster], [action], [formaction], [background], [cite], [srcdoc], [imagesrcset]";
    private const string CssCustomPropertyNamePattern = "--(?:\\\\[0-9A-Fa-f]{1,6}\\s?|\\\\[^\\r\\n\\f]|[\\p{L}\\p{N}_-]|[^\\x00-\\x7F])+";
    private static readonly Regex CssUrlExpression = new Regex("(?<name>(?:[uU]|\\\\0{0,4}(?:75|55)\\s?|\\\\[uU])(?:[rR]|\\\\0{0,4}(?:72|52)\\s?|\\\\[rR])(?:[lL]|\\\\0{0,4}(?:6[cC]|4[cC])\\s?|\\\\[lL]))\\(\\s*(?:\"(?<url>[^\"]+)\"|'(?<url>[^']+)'|(?<url>[^)]+))\\s*\\)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly Regex CssVarExpression = new Regex("(?<nameToken>(?:[vV]|\\\\0{0,4}(?:76|56)\\s?|\\\\[vV])(?:[aA]|\\\\0{0,4}(?:61|41)\\s?|\\\\[aA])(?:[rR]|\\\\0{0,4}(?:72|52)\\s?|\\\\[rR]))\\(\\s*(?<name>" + CssCustomPropertyNamePattern + ")", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly Regex CssCustomPropertyDeclarationExpression = new Regex("(?<name>" + CssCustomPropertyNamePattern + ")\\s*:", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly Regex MediaLengthFeatureExpression = new Regex("\\(\\s*(?<name>max-width|max-height|width|height)\\s*:\\s*(?<value>[+-]?(?:\\d+\\.?\\d*|\\.\\d+))\\s*(?<unit>px|em|rem|vw|vh|vmin|vmax|cm|mm|in|pt|pc)?\\s*\\)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

    /// <summary>
    /// Parses raw HTML and builds a resource manifest.
    /// </summary>
    public static HtmlResourceManifest BuildManifest(string html, HtmlResourcePipelineOptions? options = null) {
        HtmlResourcePipelineOptions resolved = options ?? new HtmlResourcePipelineOptions();
        HtmlConversionLimits limits = resolved.Limits ?? HtmlConversionLimits.CreateUntrustedProfile();
        limits.Validate();
        IHtmlDocument document = HtmlConversionDocument.ParseSourceDocumentForAnalysis(html, limits);
        return BuildManifest(document, resolved);
    }

    /// <summary>
    /// Builds a resource manifest from a parsed document.
    /// </summary>
    public static HtmlResourceManifest BuildManifest(IHtmlDocument document, HtmlResourcePipelineOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        options = options ?? new HtmlResourcePipelineOptions();
        HtmlConversionLimits limits = options.Limits ?? HtmlConversionLimits.CreateUntrustedProfile();
        limits.Validate();
        HtmlConversionInputGuard.ValidateDocument(document, limits);
        Uri? baseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(document, options.BaseUri);
        var manifest = new HtmlResourceManifest();
        foreach (IElement element in document.QuerySelectorAll(ResourceSelector)) {
            AddElementResources(manifest, element, baseUri, options, 0);
        }

        AddCssResources(manifest, document, baseUri, options);
        return manifest;
    }

    private static HtmlUrlPolicy GetResourceUrlPolicy(HtmlResourcePipelineOptions options) =>
        options.ResourceUrlPolicy ?? HtmlResourceUrlPolicy.Create(options.UrlPolicy);

}
