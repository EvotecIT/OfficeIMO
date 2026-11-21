using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

public static partial class HtmlResourcePipeline {
    private static void AddCssResources(HtmlResourceManifest manifest, IHtmlDocument document, Uri? baseUri, HtmlResourcePipelineOptions options) {
        Dictionary<string, List<CssCustomPropertyDefinition>> documentCustomPropertyDefinitions = ExtractDocumentCustomPropertyDefinitions(document, options);
        Dictionary<IElement, int> inlineSourceOrders = GetInlineStyleSourceOrders(document, GetDocumentCssSourceOrder(document));
        var computedStyles = new Lazy<IReadOnlyDictionary<IElement, HtmlComputedStyle>>(
            () => HtmlComputedStyleEngine.Compute(document, options),
            LazyThreadSafetyMode.None);
        foreach (IElement styleElement in document.QuerySelectorAll("style")) {
            if (!IsCssStyleElement(styleElement) || !IsApplicableMedia(styleElement.GetAttribute("media") ?? string.Empty, options)) {
                continue;
            }

            AddCssReferences(manifest, styleElement, "css", HtmlStyleElementText.Read(styleElement), documentCustomPropertyDefinitions, inlineSourceOrders, computedStyles, sourceOrderBase: 0, includeLocalDefinitions: false, baseUri, options, document);
        }

        foreach (IElement element in document.QuerySelectorAll("[style]")) {
            int sourceOrderBase = inlineSourceOrders.TryGetValue(element, out int inlineSourceOrder)
                ? inlineSourceOrder
                : GetDocumentCssSourceOrder(document);
            Dictionary<string, List<CssCustomPropertyDefinition>> inheritedDefinitions = ExtractInlineCustomPropertyDefinitions(element, inlineSourceOrders, options, includeSelf: false);
            Dictionary<string, List<CssCustomPropertyDefinition>> ambientDefinitions = MergeCustomPropertyDefinitions(documentCustomPropertyDefinitions, inheritedDefinitions);
            AddCssReferences(manifest, element, "style", element.GetAttribute("style") ?? string.Empty, ambientDefinitions, inlineSourceOrders, computedStyles, sourceOrderBase, includeLocalDefinitions: true, baseUri, options, document);
        }

        if (!options.DeferPageMarginBackgrounds) AddPageMarginBackgroundResources(manifest, document, baseUri, options);
    }


    private static void AddCssReferences(
        HtmlResourceManifest manifest,
        IElement element,
        string attributeName,
        string css,
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> ambientCustomPropertyDefinitions,
        IReadOnlyDictionary<IElement, int> inlineSourceOrders,
        Lazy<IReadOnlyDictionary<IElement, HtmlComputedStyle>> computedStyles,
        int sourceOrderBase,
        bool includeLocalDefinitions,
        Uri? baseUri,
        HtmlResourcePipelineOptions options,
        IHtmlDocument? document) {
        if (string.IsNullOrWhiteSpace(css)) {
            return;
        }

        css = StripCssCommentsOutsideStrings(css);
        bool requiresComputedStyles = CssUrlExpression.IsMatch(css)
            || CssVarExpression.IsMatch(css)
            || ExtractImageSetStringUrls(css).Any();
        List<SourceRange> inactiveMediaRanges = GetInactiveCssRuleRanges(css, options);
        bool scanImports = !string.Equals(attributeName, "style", StringComparison.OrdinalIgnoreCase);
        IElement? inlineUseElement = string.Equals(attributeName, "style", StringComparison.OrdinalIgnoreCase)
            ? element
            : null;
        Dictionary<string, List<CssCustomPropertyDefinition>> customPropertyDefinitions = includeLocalDefinitions
            ? MergeCustomPropertyDefinitions(ambientCustomPropertyDefinitions, ExtractCustomPropertyDefinitions(css, inactiveMediaRanges, sourceOrderBase, isInline: string.Equals(attributeName, "style", StringComparison.OrdinalIgnoreCase), sourceOwner: inlineUseElement))
            : CloneCustomPropertyDefinitions(ambientCustomPropertyDefinitions);
        var importRanges = new List<SourceRange>();
        if (scanImports) {
            foreach (CssImportReference reference in ExtractCssImports(css)) {
                string source = reference.Source;
                if (!string.IsNullOrWhiteSpace(source)
                    && !IsInRanges(reference.Start, inactiveMediaRanges)
                    && IsApplicableCssImport(reference.ConditionText, options, out _, out _)) {
                    importRanges.Add(new SourceRange(reference.Start, reference.End));
                    AddRaw(manifest, HtmlResourceKind.Stylesheet, element, attributeName + "-import", DecodeCssEscapes(source), baseUri, options);
                }
            }
        }

        if (!requiresComputedStyles) {
            return;
        }

        // Most references never consult the cascade: a font-face or ordinary rule URL needs it
        // only for custom properties or keyframes. Compute it on that first use.
        IReadOnlyDictionary<IElement, HtmlComputedStyle> resolvedComputedStyles = new DeferredComputedStyles(computedStyles);
        AddUsedCustomPropertyUrls(manifest, element, attributeName, css, customPropertyDefinitions, inlineSourceOrders, resolvedComputedStyles, inactiveMediaRanges, baseUri, options, document, inlineUseElement);
        foreach (CssStringUrlReference reference in ExtractImageSetStringUrls(css)) {
            string source = DecodeCssEscapes(reference.Source);
            if (!IsInRanges(reference.Start, inactiveMediaRanges)
                && !IsFragmentOnlyReference(source)
                && !TryGetCustomPropertyName(css, reference.Start, out _)
                && IsSupportedCssUrlDeclaration(css, reference.Start)
                && IsCssReferenceForMatchingSelector(document, attributeName, css, reference.Start, resolvedComputedStyles)) {
                AddRaw(manifest, ClassifyCssUrl(css, reference.Start), element, attributeName + "-image-set", source, baseUri, options);
            }
        }

        foreach (Match match in CssUrlExpression.Matches(css)) {
            if (!IsValidCssUrlMatch(css, match)) continue;
            string source = DecodeCssEscapes(match.Groups["url"].Value.Trim().Trim('\'', '"'));
            if (!string.IsNullOrWhiteSpace(source)
                && !IsFragmentOnlyReference(source)
                && IsCssFunctionNameAt(css, match.Index, "url")
                && !IsImportUrl(match.Index, importRanges)
                && !IsResolvedVarFallbackUrl(css, match.Index, customPropertyDefinitions, inlineSourceOrders, document, inlineUseElement, resolvedComputedStyles, inactiveMediaRanges, options, attributeName)
                && !IsInRanges(match.Index, inactiveMediaRanges)
                && !IsImportAtRuleUrl(css, match.Index)
                && !IsAtRulePreludeUrl(css, match.Index)
                && !IsInsideCssString(css, match.Index)
                && !IsCustomPropertyUrl(css, match.Index)
                && IsSupportedCssUrlDeclaration(css, match.Index)
                && IsCssReferenceForMatchingSelector(document, attributeName, css, match.Index, resolvedComputedStyles)) {
                AddRaw(manifest, ClassifyCssUrl(css, match.Index), element, attributeName + "-url", source, baseUri, options);
            }
        }
    }


    /// <summary>The document's computed styles, computed only when a reference first reads them.</summary>
    private sealed class DeferredComputedStyles : IReadOnlyDictionary<IElement, HtmlComputedStyle> {
        private readonly Lazy<IReadOnlyDictionary<IElement, HtmlComputedStyle>> _styles;

        internal DeferredComputedStyles(Lazy<IReadOnlyDictionary<IElement, HtmlComputedStyle>> styles) => _styles = styles;

        public HtmlComputedStyle this[IElement key] => _styles.Value[key];
        public IEnumerable<IElement> Keys => _styles.Value.Keys;
        public IEnumerable<HtmlComputedStyle> Values => _styles.Value.Values;
        public int Count => _styles.Value.Count;
        public bool ContainsKey(IElement key) => _styles.Value.ContainsKey(key);
        public bool TryGetValue(IElement key,
#if NETCOREAPP3_0_OR_GREATER
            [System.Diagnostics.CodeAnalysis.MaybeNullWhen(false)]
#endif
            out HtmlComputedStyle value) => _styles.Value.TryGetValue(key, out value);
        public IEnumerator<KeyValuePair<IElement, HtmlComputedStyle>> GetEnumerator() => _styles.Value.GetEnumerator();
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
