using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

public static partial class HtmlResourcePipeline {
    internal static bool HasStylesheetUrlResources(string css) {
        if (string.IsNullOrWhiteSpace(css)) {
            return false;
        }

        string normalized = StripCssCommentsOutsideStrings(css);
        var importRanges = ExtractCssImports(normalized)
            .Select(import => new SourceRange(import.Start, import.End))
            .ToList();
        foreach (Match match in CssUrlExpression.Matches(normalized)) {
            if (IsCssFunctionNameAt(normalized, match.Index, "url")
                && !IsInsideCssString(normalized, match.Index)
                && !IsImportUrl(match.Index, importRanges)) {
                return true;
            }
        }

        return ExtractImageSetStringUrls(normalized).Any(reference => !IsInRanges(reference.Start, importRanges));
    }

    private static void AddCssResources(HtmlResourceManifest manifest, IHtmlDocument document, Uri? baseUri, HtmlResourcePipelineOptions options) {
        Dictionary<string, List<CssCustomPropertyDefinition>> documentCustomPropertyDefinitions = ExtractDocumentCustomPropertyDefinitions(document, options.MediaContext);
        Dictionary<IElement, int> inlineSourceOrders = GetInlineStyleSourceOrders(document, GetDocumentCssSourceOrder(document));
        foreach (IElement styleElement in document.QuerySelectorAll("style")) {
            if (!IsCssStyleElement(styleElement) || !IsApplicableMedia(styleElement.GetAttribute("media") ?? string.Empty, options.MediaContext)) {
                continue;
            }

            AddCssReferences(manifest, styleElement, "css", styleElement.TextContent, documentCustomPropertyDefinitions, inlineSourceOrders, sourceOrderBase: 0, includeLocalDefinitions: false, baseUri, options, document);
        }

        foreach (IElement element in document.QuerySelectorAll("[style]")) {
            int sourceOrderBase = inlineSourceOrders.TryGetValue(element, out int inlineSourceOrder)
                ? inlineSourceOrder
                : GetDocumentCssSourceOrder(document);
            Dictionary<string, List<CssCustomPropertyDefinition>> inheritedDefinitions = ExtractInlineCustomPropertyDefinitions(element, inlineSourceOrders, options.MediaContext, includeSelf: false);
            Dictionary<string, List<CssCustomPropertyDefinition>> ambientDefinitions = MergeCustomPropertyDefinitions(documentCustomPropertyDefinitions, inheritedDefinitions);
            AddCssReferences(manifest, element, "style", element.GetAttribute("style") ?? string.Empty, ambientDefinitions, inlineSourceOrders, sourceOrderBase, includeLocalDefinitions: true, baseUri, options, document);
        }
    }


    private static void AddCssReferences(
        HtmlResourceManifest manifest,
        IElement element,
        string attributeName,
        string css,
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> ambientCustomPropertyDefinitions,
        IReadOnlyDictionary<IElement, int> inlineSourceOrders,
        int sourceOrderBase,
        bool includeLocalDefinitions,
        Uri? baseUri,
        HtmlResourcePipelineOptions options,
        IHtmlDocument? document) {
        if (string.IsNullOrWhiteSpace(css)) {
            return;
        }

        css = StripCssCommentsOutsideStrings(css);
        List<SourceRange> inactiveMediaRanges = GetInactiveCssRuleRanges(css, options.MediaContext);
        bool scanImports = !string.Equals(attributeName, "style", StringComparison.OrdinalIgnoreCase);
        IElement? inlineUseElement = string.Equals(attributeName, "style", StringComparison.OrdinalIgnoreCase)
            ? element
            : null;
        Dictionary<string, List<CssCustomPropertyDefinition>> customPropertyDefinitions = includeLocalDefinitions
            ? MergeCustomPropertyDefinitions(ambientCustomPropertyDefinitions, ExtractCustomPropertyDefinitions(css, inactiveMediaRanges, sourceOrderBase, isInline: string.Equals(attributeName, "style", StringComparison.OrdinalIgnoreCase), inlineOwner: inlineUseElement))
            : CloneCustomPropertyDefinitions(ambientCustomPropertyDefinitions);
        var importRanges = new List<SourceRange>();
        if (scanImports) {
            foreach (CssImportReference reference in ExtractCssImports(css)) {
                string source = reference.Source;
                if (!string.IsNullOrWhiteSpace(source)
                    && !IsInRanges(reference.Start, inactiveMediaRanges)
                    && IsApplicableCssImport(reference.ConditionText, options.MediaContext)) {
                    importRanges.Add(new SourceRange(reference.Start, reference.End));
                    AddRaw(manifest, HtmlResourceKind.Stylesheet, element, attributeName + "-import", DecodeCssEscapes(source), baseUri, options);
                }
            }
        }

        AddUsedCustomPropertyUrls(manifest, element, attributeName, css, customPropertyDefinitions, inlineSourceOrders, inactiveMediaRanges, baseUri, options, document, inlineUseElement);
        foreach (CssStringUrlReference reference in ExtractImageSetStringUrls(css)) {
            string source = DecodeCssEscapes(reference.Source);
            if (!IsInRanges(reference.Start, inactiveMediaRanges)
                && !IsFragmentOnlyReference(source)
                && !TryGetCustomPropertyName(css, reference.Start, out _)
                && IsSupportedCssUrlDeclaration(css, reference.Start)
                && IsCssReferenceForMatchingSelector(document, attributeName, css, reference.Start)) {
                AddRaw(manifest, ClassifyCssUrl(css, reference.Start), element, attributeName + "-image-set", source, baseUri, options);
            }
        }

        foreach (Match match in CssUrlExpression.Matches(css)) {
            string source = DecodeCssEscapes(match.Groups["url"].Value.Trim().Trim('\'', '"'));
            if (!string.IsNullOrWhiteSpace(source)
                && !IsFragmentOnlyReference(source)
                && IsCssFunctionNameAt(css, match.Index, "url")
                && !IsImportUrl(match.Index, importRanges)
                && !IsResolvedVarFallbackUrl(css, match.Index, customPropertyDefinitions, inlineSourceOrders, document, inlineUseElement, inactiveMediaRanges, options, attributeName)
                && !IsInRanges(match.Index, inactiveMediaRanges)
                && !IsImportAtRuleUrl(css, match.Index)
                && !IsAtRulePreludeUrl(css, match.Index)
                && !IsInsideCssString(css, match.Index)
                && !IsCustomPropertyUrl(css, match.Index)
                && IsSupportedCssUrlDeclaration(css, match.Index)
                && IsCssReferenceForMatchingSelector(document, attributeName, css, match.Index)) {
                AddRaw(manifest, ClassifyCssUrl(css, match.Index), element, attributeName + "-url", source, baseUri, options);
            }
        }
    }

}
