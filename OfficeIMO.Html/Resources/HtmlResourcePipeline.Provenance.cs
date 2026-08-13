using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

public static partial class HtmlResourcePipeline {
    internal static bool IsActiveProvenanceStyleElement(IElement element) =>
        IsCssStyleElement(element) && IsApplicableMedia(element.GetAttribute("media") ?? string.Empty, new HtmlResourcePipelineOptions());

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
        Dictionary<string, List<CssCustomPropertyDefinition>> documentDefinitions = ExtractDocumentCustomPropertyDefinitions(document, options);
        Dictionary<IElement, int> inlineSourceOrders = GetInlineStyleSourceOrders(document, GetDocumentCssSourceOrder(document));
        var result = new HtmlProvenanceCssScope();

        foreach (IElement styleElement in document.QuerySelectorAll("style")) {
            if (!IsCssStyleElement(styleElement) || !IsApplicableMedia(styleElement.GetAttribute("media") ?? string.Empty, options)) continue;
            CollectResolvedCustomPropertyDeclarations(styleElement.TextContent, documentDefinitions, inlineSourceOrders, document, null, options, result.UsedCustomPropertyDeclarations);
            CollectResolvedVarFallbackStarts(styleElement.TextContent, documentDefinitions, inlineSourceOrders, document, null, options, "css", styleElement, result);
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
            CollectResolvedCustomPropertyDeclarations(css, definitions, inlineSourceOrders, document, element, options, result.UsedCustomPropertyDeclarations);
            CollectResolvedVarFallbackStarts(css, definitions, inlineSourceOrders, document, element, options, "style", element, result);
        }
        return result;
    }

    private static void CollectResolvedVarFallbackStarts(
        string css,
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> definitions,
        IReadOnlyDictionary<IElement, int> inlineSourceOrders,
        IHtmlDocument document,
        IElement? inlineUseElement,
        HtmlResourcePipelineOptions options,
        string attributeName,
        IElement sourceOwner,
        HtmlProvenanceCssScope result) {
        if (string.IsNullOrWhiteSpace(css) || definitions.Count == 0) return;
        string masked = MaskCssComments(css);
        List<SourceRange> inactiveRanges = GetInactiveCssRuleRanges(masked, options);
        foreach (Match match in CssUrlExpression.Matches(masked)) {
            if (IsResolvedVarFallbackUrl(masked, match.Index, definitions, inlineSourceOrders, document,
                inlineUseElement, inactiveRanges, options, attributeName)) {
                result.AddResolvedFallback(sourceOwner, match.Index);
            }
        }
        foreach (CssStringUrlReference reference in ExtractImageSetStringUrls(masked)) {
            if (IsResolvedVarFallbackUrl(masked, reference.Start, definitions, inlineSourceOrders, document,
                inlineUseElement, inactiveRanges, options, attributeName)) {
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
                    propertyName, effectiveDefinitions, useSelector, document, useElement,
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
        ISet<int>? resolvedVarFallbackStarts = null) {
        if (string.IsNullOrWhiteSpace(css)) yield break;
        string masked = MaskCssComments(css);
        List<SourceRange> inactiveRanges = GetInactiveCssRuleRanges(masked, new HtmlResourcePipelineOptions());
        var emittedRanges = new HashSet<(int Start, int Length)>();
        foreach (Match match in CssUrlExpression.Matches(masked)) {
            bool isCustomProperty = TryGetCustomPropertyName(masked, match.Index, out _);
            if (IsInRanges(match.Index, inactiveRanges) || resolvedVarFallbackStarts?.Contains(match.Index) == true ||
                !IsCssFunctionNameAt(masked, match.Index, "url") ||
                IsInsideCssString(masked, match.Index) ||
                IsImportAtRuleUrl(masked, match.Index) ||
                IsAtRulePreludeUrl(masked, match.Index) ||
                !IsCssReferenceForMatchingSelector(document, attributeName, masked, match.Index) ||
                isCustomProperty && (!TryGetCustomPropertyDeclarationStart(masked, match.Index, out int declarationStart) ||
                    usedCustomPropertyDeclarationStarts == null || !usedCustomPropertyDeclarationStarts.Contains(declarationStart)) ||
                !isCustomProperty && ClassifyCssUrl(masked, match.Index) != HtmlResourceKind.Image) continue;
            Group sourceGroup = match.Groups["url"];
            int leading = 0;
            while (leading < sourceGroup.Length && char.IsWhiteSpace(sourceGroup.Value[leading])) leading++;
            int trailing = sourceGroup.Length;
            while (trailing > leading && char.IsWhiteSpace(sourceGroup.Value[trailing - 1])) trailing--;
            if (trailing == leading) continue;
            string source = DecodeCssEscapes(sourceGroup.Value.Substring(leading, trailing - leading));
            var range = (sourceGroup.Index + leading, trailing - leading);
            if (emittedRanges.Add(range)) yield return new HtmlCssImageReference(range.Item1, range.Item2, source);
        }

        foreach (CssStringUrlReference reference in ExtractImageSetStringUrls(masked)) {
            bool isCustomProperty = TryGetCustomPropertyName(masked, reference.Start, out _);
            if (IsInRanges(reference.Start, inactiveRanges) || resolvedVarFallbackStarts?.Contains(reference.Start) == true ||
                !IsCssReferenceForMatchingSelector(document, attributeName, masked, reference.Start) ||
                (isCustomProperty
                    ? !TryGetCustomPropertyDeclarationStart(masked, reference.Start, out int declarationStart) ||
                        usedCustomPropertyDeclarationStarts == null || !usedCustomPropertyDeclarationStarts.Contains(declarationStart)
                    : ClassifyCssUrl(masked, reference.Start) != HtmlResourceKind.Image)) continue;
            if (!emittedRanges.Add((reference.SourceStart, reference.Source.Length))) continue;
            yield return new HtmlCssImageReference(reference.SourceStart, reference.Source.Length, DecodeCssEscapes(reference.Source));
        }
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
            result[index++] = ' ';
            result[index] = ' ';
            while (index + 1 < css.Length && !(css[index] == '*' && css[index + 1] == '/')) result[++index] = ' ';
            if (index + 1 < css.Length) { result[index] = ' '; result[++index] = ' '; }
        }
        return result.ToString();
    }
}

internal sealed class HtmlProvenanceCssScope {
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
