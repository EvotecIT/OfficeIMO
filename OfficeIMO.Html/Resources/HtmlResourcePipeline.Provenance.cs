using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

public static partial class HtmlResourcePipeline {
    private const char CssCommentMask = '\u0001';
    internal static bool IsActiveProvenanceStyleElement(IElement element) =>
        IsCssStyleElement(element) && IsApplicableMedia(element.GetAttribute("media") ?? string.Empty, new HtmlResourcePipelineOptions());

    internal static bool IsApplicableProvenanceMedia(IElement element) =>
        IsApplicableMedia(element.GetAttribute("media") ?? string.Empty, new HtmlResourcePipelineOptions());

    internal static bool IsActivePictureImageSource(IElement element) {
        var options = new HtmlResourcePipelineOptions();
        if (!string.Equals(element.ParentElement?.LocalName, "picture", StringComparison.OrdinalIgnoreCase)) return true;
        return IsFirstApplicablePictureSource(element, baseUri: null, options) &&
            HasPictureSourceCandidate(element) &&
            IsApplicableMedia(element.GetAttribute("media") ?? string.Empty, options) &&
            IsSupportedPictureSourceType(element.GetAttribute("type"));
    }

    internal static bool IsActivePictureFallbackImage(IElement element) =>
        !HasSelectedPictureSourceBeforeFallback(element, baseUri: null, new HtmlResourcePipelineOptions());

    internal static HtmlProvenanceCssScope CollectProvenanceCssImageScope(IHtmlDocument document) {
        var options = new HtmlResourcePipelineOptions();
        HtmlComputedStyleSet computedStyleSet = HtmlComputedStyleEngine.ComputeForProvenance(document, options);
        IReadOnlyDictionary<IElement, HtmlComputedStyle> computedStyles = computedStyleSet.Elements;
        Dictionary<string, List<CssCustomPropertyDefinition>> documentDefinitions = ExtractDocumentCustomPropertyDefinitions(document, options);
        Dictionary<IElement, int> inlineSourceOrders = GetInlineStyleSourceOrders(document, GetDocumentCssSourceOrder(document));
        var result = new HtmlProvenanceCssScope(computedStyleSet);

        foreach (IElement styleElement in document.QuerySelectorAll("style")) {
            if (!IsCssStyleElement(styleElement) || !IsApplicableMedia(styleElement.GetAttribute("media") ?? string.Empty, options)) continue;
            CollectResolvedCustomPropertyDeclarations(styleElement.TextContent, documentDefinitions, inlineSourceOrders, document, null, computedStyles, options, result.UsedCustomPropertyDeclarations);
            CollectResolvedVarFallbackStarts(styleElement.TextContent, documentDefinitions, inlineSourceOrders, document, null, computedStyles, options, "css", styleElement, result);
        }
        foreach (IElement element in document.QuerySelectorAll("[style]")) {
            string css = element.GetAttribute("style") ?? string.Empty;
            int sourceOrderBase = inlineSourceOrders.TryGetValue(element, out int sourceOrder)
                ? sourceOrder
                : GetDocumentCssSourceOrder(document);
            List<SourceRange> inactiveRanges = GetInactiveCssRuleRanges(css, options);
            Dictionary<string, List<CssCustomPropertyDefinition>> definitions = MergeCustomPropertyDefinitions(
                documentDefinitions,
                ExtractInlineCustomPropertyDefinitions(element, inlineSourceOrders, options, includeSelf: false));
            definitions = MergeCustomPropertyDefinitions(definitions,
                ExtractCustomPropertyDefinitions(css, inactiveRanges, sourceOrderBase, isInline: true, sourceOwner: element));
            CollectResolvedCustomPropertyDeclarations(css, definitions, inlineSourceOrders, document, element, computedStyles, options, result.UsedCustomPropertyDeclarations);
            CollectResolvedVarFallbackStarts(css, definitions, inlineSourceOrders, document, element, computedStyles, options, "style", element, result);
        }
        return result;
    }

    private static void CollectResolvedVarFallbackStarts(
        string css,
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> definitions,
        IReadOnlyDictionary<IElement, int> inlineSourceOrders,
        IHtmlDocument document,
        IElement? inlineUseElement,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> computedStyles,
        HtmlResourcePipelineOptions options,
        string attributeName,
        IElement sourceOwner,
        HtmlProvenanceCssScope result) {
        if (string.IsNullOrWhiteSpace(css) || definitions.Count == 0) return;
        string masked = MaskCssComments(css);
        List<SourceRange> inactiveRanges = GetInactiveCssRuleRanges(masked, options);
        foreach (Match match in CssUrlExpression.Matches(masked)) {
            if (!IsValidCssUrlMatch(masked, match)) continue;
            if (IsResolvedVarFallbackUrl(masked, match.Index, definitions, inlineSourceOrders, document,
                inlineUseElement, computedStyles, inactiveRanges, options, attributeName)) {
                result.AddResolvedFallback(sourceOwner, match.Index);
            }
        }
        foreach (CssStringUrlReference reference in ExtractImageSetStringUrls(masked)) {
            if (IsResolvedVarFallbackUrl(masked, reference.Start, definitions, inlineSourceOrders, document,
                inlineUseElement, computedStyles, inactiveRanges, options, attributeName)) {
                result.AddResolvedFallback(sourceOwner, reference.Start);
            }
        }
    }

    private static void CollectResolvedCustomPropertyDeclarations(
        string css,
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> definitions,
        IReadOnlyDictionary<IElement, int> inlineSourceOrders,
        IHtmlDocument document,
        IElement? inlineUseElement,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> computedStyles,
        HtmlResourcePipelineOptions options,
        IDictionary<IElement, HashSet<int>> result) {
        if (string.IsNullOrWhiteSpace(css) || definitions.Count == 0) return;
        string masked = MaskCssComments(css);
        List<SourceRange> inactiveRanges = GetInactiveCssRuleRanges(masked, options);
        foreach (Match variable in CssVarExpression.Matches(masked)) {
            if (IsInRanges(variable.Index, inactiveRanges) ||
                !IsCssFunctionNameAt(masked, variable.Index, "var") || IsInsideCssString(masked, variable.Index) ||
                ClassifyCssUrl(masked, variable.Index) != HtmlResourceKind.Image) continue;
            string propertyName = DecodeCssEscapes(variable.Groups["name"].Value);
            string useSelector = GetDeclarationSelector(masked, variable.Index);
            IEnumerable<IElement?> useElements = inlineUseElement != null
                ? new IElement?[] { inlineUseElement }
                : GetElementsMatchingSelectorList(document, useSelector).Cast<IElement?>();
            foreach (IElement? useElement in useElements) {
                IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> effectiveDefinitions = definitions;
                if (inlineUseElement == null && useElement != null) {
                    Dictionary<string, List<CssCustomPropertyDefinition>> inlineDefinitions =
                        ExtractInlineCustomPropertyDefinitions(useElement, inlineSourceOrders, options, includeSelf: true);
                    if (inlineDefinitions.Count != 0) effectiveDefinitions = MergeCustomPropertyDefinitions(definitions, inlineDefinitions);
                }
                foreach (CssCustomPropertyDefinition source in ResolveCustomPropertyUrlDefinitions(
                    propertyName, effectiveDefinitions, useSelector, document, useElement, computedStyles,
                    new HashSet<string>(StringComparer.Ordinal), depth: 0)) {
                    if (source.SourceOwner == null) continue;
                    if (!result.TryGetValue(source.SourceOwner, out HashSet<int>? starts)) {
                        starts = new HashSet<int>();
                        result.Add(source.SourceOwner, starts);
                    }
                    starts.Add(source.LocalDeclarationStart);
                }
            }
        }
    }

    internal static IEnumerable<HtmlCssImageReference> EnumerateProvenanceCssImageReferences(
        IHtmlDocument document,
        string attributeName,
        string css,
        ISet<int>? usedCustomPropertyDeclarationStarts = null,
        ISet<int>? resolvedVarFallbackStarts = null,
        IReadOnlyDictionary<IElement, HtmlComputedStyle>? computedStyles = null,
        HtmlComputedStyleSet? computedStyleSet = null,
        IElement? sourceOwner = null) {
        if (string.IsNullOrWhiteSpace(css)) yield break;
        string masked = MaskCssComments(css);
        List<SourceRange> inactiveRanges = GetInactiveCssRuleRanges(masked, new HtmlResourcePipelineOptions());
        var emittedRanges = new HashSet<(int Start, int Length)>();
        foreach (Match match in CssUrlExpression.Matches(masked)) {
            if (!IsValidCssUrlMatch(masked, match)) continue;
            bool isCustomProperty = TryGetCustomPropertyName(masked, match.Index, out _);
            if (IsInRanges(match.Index, inactiveRanges) || resolvedVarFallbackStarts?.Contains(match.Index) == true ||
                !IsCssFunctionNameAt(masked, match.Index, "url") ||
                IsInsideCssString(masked, match.Index) ||
                IsImportAtRuleUrl(masked, match.Index) ||
                IsAtRulePreludeUrl(masked, match.Index) ||
                !IsCssReferenceForMatchingSelector(document, attributeName, masked, match.Index, computedStyles) ||
                isCustomProperty && (!TryGetCustomPropertyDeclarationStart(masked, match.Index, out int declarationStart) ||
                    usedCustomPropertyDeclarationStarts == null || !usedCustomPropertyDeclarationStarts.Contains(declarationStart)) ||
                !isCustomProperty && ClassifyCssUrl(masked, match.Index) != HtmlResourceKind.Image) continue;
            Group sourceGroup = match.Groups["url"];
            int leading = 0;
            while (leading < sourceGroup.Length && IsCssWhitespace(sourceGroup.Value[leading])) leading++;
            int trailing = sourceGroup.Length;
            while (trailing > leading && IsCssWhitespace(sourceGroup.Value[trailing - 1])) trailing--;
            if (trailing == leading) continue;
            string source = DecodeCssEscapes(sourceGroup.Value.Substring(leading, trailing - leading));
            if (!isCustomProperty && !IsEffectiveImageDeclaration(
                    document, attributeName, sourceOwner, masked, match.Index, source, computedStyles, computedStyleSet)) continue;
            var range = (sourceGroup.Index + leading, trailing - leading);
            if (emittedRanges.Add(range)) yield return new HtmlCssImageReference(range.Item1, range.Item2, source);
        }

        foreach (CssStringUrlReference reference in ExtractImageSetStringUrls(masked)) {
            bool isCustomProperty = TryGetCustomPropertyName(masked, reference.Start, out _);
            if (IsInRanges(reference.Start, inactiveRanges) || resolvedVarFallbackStarts?.Contains(reference.Start) == true ||
                !IsCssReferenceForMatchingSelector(document, attributeName, masked, reference.Start, computedStyles) ||
                (isCustomProperty
                    ? !TryGetCustomPropertyDeclarationStart(masked, reference.Start, out int declarationStart) ||
                        usedCustomPropertyDeclarationStarts == null || !usedCustomPropertyDeclarationStarts.Contains(declarationStart)
                    : ClassifyCssUrl(masked, reference.Start) != HtmlResourceKind.Image)) continue;
            if (!isCustomProperty && !IsEffectiveImageDeclaration(
                    document, attributeName, sourceOwner, masked, reference.Start,
                    DecodeCssEscapes(reference.Source), computedStyles, computedStyleSet)) continue;
            if (!emittedRanges.Add((reference.SourceStart, reference.Source.Length))) continue;
            yield return new HtmlCssImageReference(reference.SourceStart, reference.Source.Length, DecodeCssEscapes(reference.Source));
        }
    }

    private static bool IsEffectiveImageDeclaration(
        IHtmlDocument document,
        string attributeName,
        IElement? sourceOwner,
        string css,
        int index,
        string source,
        IReadOnlyDictionary<IElement, HtmlComputedStyle>? computedStyles,
        HtmlComputedStyleSet? computedStyleSet) {
        if (TryGetEnclosingKeyframesName(css, index, out _, out _)) return true;
        string propertyName = GetCssDeclarationPropertyName(css, index);
        if (propertyName.Length == 0) return true;
        int declarationStart = GetDeclarationStart(css, index);
        int colon = css.IndexOf(':', declarationStart, Math.Max(0, index - declarationStart));
        if (colon < 0) return true;
        int valueEnd = FindDeclarationValueEnd(css, colon + 1);
        string declarationValue = css.Substring(colon + 1, valueEnd - colon - 1).Trim();
        int important = declarationValue.LastIndexOf("!important", StringComparison.OrdinalIgnoreCase);
        if (important >= 0 && string.IsNullOrWhiteSpace(declarationValue.Substring(important + 10))) {
            declarationValue = declarationValue.Substring(0, important).TrimEnd();
        }
        computedStyles ??= computedStyleSet?.Elements ?? HtmlComputedStyleEngine.Compute(document);
        IEnumerable<IElement> elements;
        string selector = string.Empty;
        if (string.Equals(attributeName, "style", StringComparison.OrdinalIgnoreCase)) {
            if (sourceOwner == null) return true;
            elements = new[] { sourceOwner };
        } else {
            selector = GetDeclarationSelector(css, index);
            elements = GetElementsMatchingSelectorList(document, selector);
        }
        bool sawElement = false;
        foreach (IElement element in elements) {
            sawElement = true;
            if (!computedStyles.TryGetValue(element, out HtmlComputedStyle? elementStyle) ||
                IsDisplayNone(elementStyle)) continue;
            foreach (HtmlComputedStyle style in GetDeclarationStyles(element, elementStyle, selector, computedStyleSet)) {
                if (IsDisplayNone(style)) continue;
                string effective = style.GetValue(propertyName);
                if (effective.Length == 0) {
                    if (!IsInsideContainerRule(css, index)) return true;
                    continue;
                }
                if (ContainsEquivalentImageSource(effective, source) || string.Equals(
                        HtmlRenderCssValues.NormalizeComponentValueWhitespace(DecodeCssEscapes(declarationValue)),
                        HtmlRenderCssValues.NormalizeComponentValueWhitespace(effective),
                        StringComparison.Ordinal)) return true;
            }
        }
        return !sawElement;
    }

    private static IEnumerable<HtmlComputedStyle> GetDeclarationStyles(
        IElement element,
        HtmlComputedStyle elementStyle,
        string selector,
        HtmlComputedStyleSet? computedStyleSet) {
        bool yieldedPseudo = false;
        if (computedStyleSet != null && selector.Length != 0) {
            foreach (string selectorPart in SplitTopLevelList(selector)) {
                if (!HtmlComputedStyleEngine.TryParsePseudoElementSelector(
                        selectorPart, out string hostSelector, out HtmlPseudoElementKind kind)) continue;
                string normalized = NormalizeSelectorForQuery(hostSelector, stripStatefulPseudoClasses: true);
                if (normalized.Length == 0) normalized = "*";
                bool matches;
                try { matches = element.Matches(normalized); } catch { matches = false; }
                if (!matches || !computedStyleSet.TryGetPseudoStyle(element, kind, out HtmlComputedStyle pseudoStyle)) continue;
                yieldedPseudo = true;
                yield return pseudoStyle;
            }
        }
        if (!yieldedPseudo) yield return elementStyle;
    }

    private static bool IsDisplayNone(HtmlComputedStyle style) =>
        string.Equals(style.GetValue("display").Trim(), "none", StringComparison.OrdinalIgnoreCase);

    private static bool ContainsEquivalentImageSource(string effective, string source) {
        foreach (Match match in CssUrlExpression.Matches(effective)) {
            if (!IsValidCssUrlMatch(effective, match)) continue;
            string candidate = DecodeCssEscapes(match.Groups["url"].Value.Trim());
            if (string.Equals(candidate, source, StringComparison.Ordinal)) return true;
        }
        foreach (CssStringUrlReference reference in ExtractImageSetStringUrls(effective)) {
            if (string.Equals(DecodeCssEscapes(reference.Source), source, StringComparison.Ordinal)) return true;
        }
        return false;
    }

    private static bool IsInsideContainerRule(string css, int index) {
        int search = 0;
        while (search < index) {
            int start = css.IndexOf("@container", search, StringComparison.OrdinalIgnoreCase);
            if (start < 0 || start >= index) return false;
            if (IsInsideCssString(css, start) || !HasAtRuleTokenBoundary(css, start, "@container")) {
                search = start + 10;
                continue;
            }
            int open = FindNextTopLevelBlockStart(css, start + 10);
            if (open < 0) return false;
            int close = FindMatchingCssBrace(css, open);
            if (close < 0) return false;
            if (index > open && index < close) return true;
            search = close + 1;
        }
        return false;
    }

    private static bool TryGetCustomPropertyDeclarationStart(string css, int valueIndex, out int declarationStart) {
        declarationStart = -1;
        foreach (Match declaration in CssCustomPropertyDeclarationExpression.Matches(css)) {
            if (declaration.Index > valueIndex) break;
            int colon = css.IndexOf(':', declaration.Index);
            if (colon < 0) continue;
            int end = FindDeclarationValueEnd(css, colon + 1);
            if (valueIndex >= colon + 1 && valueIndex < end) declarationStart = declaration.Index;
        }
        return declarationStart >= 0;
    }

    private static string MaskCssComments(string css) {
        var result = new StringBuilder(css);
        char quote = '\0';
        for (int index = 0; index < css.Length; index++) {
            char current = css[index];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(css, index)) quote = '\0';
                continue;
            }
            if (current is '"' or '\'') { quote = current; continue; }
            if (current != '/' || index + 1 >= css.Length || css[index + 1] != '*') continue;
            result[index++] = CssCommentMask;
            result[index] = CssCommentMask;
            while (index + 1 < css.Length && !(css[index] == '*' && css[index + 1] == '/')) result[++index] = CssCommentMask;
            if (index + 1 < css.Length) { result[index] = CssCommentMask; result[++index] = CssCommentMask; }
        }
        return result.ToString();
    }
}

internal sealed class HtmlProvenanceCssScope {
    internal HtmlProvenanceCssScope(HtmlComputedStyleSet computedStyleSet) {
        ComputedStyleSet = computedStyleSet;
    }

    internal HtmlComputedStyleSet ComputedStyleSet { get; }
    internal IReadOnlyDictionary<IElement, HtmlComputedStyle> ComputedStyles => ComputedStyleSet.Elements;
    internal Dictionary<IElement, HashSet<int>> UsedCustomPropertyDeclarations { get; } = new();
    internal Dictionary<IElement, HashSet<int>> ResolvedVarFallbackStarts { get; } = new();

    internal void AddResolvedFallback(IElement owner, int start) {
        if (!ResolvedVarFallbackStarts.TryGetValue(owner, out HashSet<int>? starts)) {
            starts = new HashSet<int>();
            ResolvedVarFallbackStarts.Add(owner, starts);
        }
        starts.Add(start);
    }
}

internal readonly struct HtmlCssImageReference {
    internal HtmlCssImageReference(int start, int length, string value) {
        Start = start;
        Length = length;
        Value = value;
    }

    internal int Start { get; }
    internal int Length { get; }
    internal string Value { get; }
}
