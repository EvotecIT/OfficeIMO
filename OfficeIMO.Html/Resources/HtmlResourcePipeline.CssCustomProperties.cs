using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

public static partial class HtmlResourcePipeline {
    private static Dictionary<string, List<CssCustomPropertyDefinition>> ExtractDocumentCustomPropertyDefinitions(IHtmlDocument document, HtmlResourcePipelineOptions options) {
        var definitions = new Dictionary<string, List<CssCustomPropertyDefinition>>(StringComparer.Ordinal);
        int sourceOrderBase = 0;
        foreach (IElement styleElement in document.QuerySelectorAll("style")) {
            string css = styleElement.TextContent;
            if (!IsCssStyleElement(styleElement) || !IsApplicableMedia(styleElement.GetAttribute("media") ?? string.Empty, options) || string.IsNullOrWhiteSpace(css)) {
                sourceOrderBase += css.Length + 1;
                continue;
            }

            css = StripCssCommentsOutsideStrings(css);
            MergeCustomPropertyDefinitionsInto(definitions, ExtractCustomPropertyDefinitions(css, GetInactiveCssRuleRanges(css, options), sourceOrderBase, isInline: false, inlineOwner: null));
            sourceOrderBase += css.Length + 1;
        }

        return definitions;
    }

    private static int GetDocumentCssSourceOrder(IHtmlDocument document) {
        int sourceOrder = 0;
        foreach (IElement styleElement in document.QuerySelectorAll("style")) {
            sourceOrder += styleElement.TextContent.Length + 1;
        }

        return sourceOrder;
    }

    private static Dictionary<IElement, int> GetInlineStyleSourceOrders(IHtmlDocument document, int sourceOrderBase) {
        var sourceOrders = new Dictionary<IElement, int>();
        int sourceOrder = sourceOrderBase;
        foreach (IElement element in document.QuerySelectorAll("[style]")) {
            sourceOrders[element] = sourceOrder;
            sourceOrder += (element.GetAttribute("style") ?? string.Empty).Length + 1;
        }

        return sourceOrders;
    }

    private static Dictionary<string, List<CssCustomPropertyDefinition>> ExtractInlineCustomPropertyDefinitions(IElement element, IReadOnlyDictionary<IElement, int> inlineSourceOrders, HtmlResourcePipelineOptions options, bool includeSelf) {
        var definitions = new Dictionary<string, List<CssCustomPropertyDefinition>>(StringComparer.Ordinal);
        for (IElement? current = includeSelf ? element : element.ParentElement; current != null; current = current.ParentElement) {
            string style = current.GetAttribute("style") ?? string.Empty;
            if (style.Length == 0 || !inlineSourceOrders.TryGetValue(current, out int sourceOrderBase)) {
                continue;
            }

            string css = StripCssCommentsOutsideStrings(style);
            MergeCustomPropertyDefinitionsInto(definitions, ExtractCustomPropertyDefinitions(css, GetInactiveCssRuleRanges(css, options), sourceOrderBase, isInline: true, inlineOwner: current));
        }

        return definitions;
    }

    private static Dictionary<string, List<CssCustomPropertyDefinition>> ExtractCustomPropertyDefinitions(string css, IReadOnlyList<SourceRange> inactiveMediaRanges, int sourceOrderBase, bool isInline, IElement? inlineOwner) {
        var definitions = new Dictionary<string, List<CssCustomPropertyDefinition>>(StringComparer.Ordinal);
        foreach (Match match in CssCustomPropertyDeclarationExpression.Matches(css)) {
            string propertyName = DecodeCssEscapes(match.Groups["name"].Value);
            int declarationStart = match.Index;
            int valueStart = css.IndexOf(':', declarationStart);
            if (IsInsideCssString(css, declarationStart)
                || IsInRanges(declarationStart, inactiveMediaRanges)
                || valueStart < 0
                || GetCssDeclarationPropertyName(css, valueStart + 1) != propertyName) {
                continue;
            }

            int valueEnd = FindDeclarationValueEnd(css, valueStart + 1);
            string selector = GetDeclarationSelector(css, declarationStart);
            bool isImportant = IsImportantDeclarationValue(css, valueStart + 1, valueEnd);
            string valueText = GetCustomPropertyValueText(css, valueStart + 1, valueEnd);
            List<string> aliases = ExtractCustomPropertyAliases(css, valueStart + 1, valueEnd);
            bool addedUrl = false;
            foreach (Match urlMatch in CssUrlExpression.Matches(css)) {
                if (urlMatch.Index < valueStart || urlMatch.Index >= valueEnd || !IsCssFunctionNameAt(css, urlMatch.Index, "url") || IsInsideCssString(css, urlMatch.Index)) {
                    continue;
                }

                string? fallbackAlias = TryGetVarFallbackAlias(css, valueStart + 1, valueEnd, urlMatch.Index);
                AddCustomPropertyDefinition(definitions, propertyName, DecodeCssEscapes(urlMatch.Groups["url"].Value.Trim().Trim('\'', '"')), selector, sourceOrderBase + declarationStart, isImportant, aliases, isInline, inlineOwner, valueText, fallbackAlias);
                addedUrl = true;
            }

            foreach (CssStringUrlReference reference in ExtractImageSetStringUrls(css)) {
                if (reference.Start < valueStart || reference.Start >= valueEnd) {
                    continue;
                }

                string? fallbackAlias = TryGetVarFallbackAlias(css, valueStart + 1, valueEnd, reference.Start);
                AddCustomPropertyDefinition(definitions, propertyName, DecodeCssEscapes(reference.Source), selector, sourceOrderBase + declarationStart, isImportant, aliases, isInline, inlineOwner, valueText, fallbackAlias);
                addedUrl = true;
            }

            if (!addedUrl) {
                AddCustomPropertyDefinition(definitions, propertyName, string.Empty, selector, sourceOrderBase + declarationStart, isImportant, aliases, isInline, inlineOwner, valueText, fallbackAlias: null);
            }
        }

        return definitions;
    }

    private static List<string> ExtractCustomPropertyAliases(string css, int valueStart, int valueEnd) {
        var aliases = new List<string>();
        foreach (Match varMatch in CssVarExpression.Matches(css)) {
            if (varMatch.Index < valueStart
                || varMatch.Index >= valueEnd
                || !IsCssFunctionNameAt(css, varMatch.Index, "var")
                || IsInsideCssString(css, varMatch.Index)) {
                continue;
            }

            string alias = DecodeCssEscapes(varMatch.Groups["name"].Value);
            if (!aliases.Contains(alias, StringComparer.Ordinal)) {
                aliases.Add(alias);
            }
        }

        return aliases;
    }

    private static string GetCustomPropertyValueText(string css, int valueStart, int valueEnd) {
        string value = css.Substring(valueStart, Math.Max(0, valueEnd - valueStart)).Trim();
        int important = value.LastIndexOf("!important", StringComparison.OrdinalIgnoreCase);
        if (important >= 0 && string.IsNullOrWhiteSpace(value.Substring(important + 10))) {
            value = value.Substring(0, important).TrimEnd();
        }

        return DecodeCssEscapes(value).Trim();
    }

    private static string? TryGetVarFallbackAlias(string css, int valueStart, int valueEnd, int urlIndex) {
        foreach (Match varMatch in CssVarExpression.Matches(css)) {
            if (varMatch.Index < valueStart
                || varMatch.Index >= valueEnd
                || urlIndex <= varMatch.Index
                || !IsCssFunctionNameAt(css, varMatch.Index, "var")
                || IsInsideCssString(css, varMatch.Index)) {
                continue;
            }

            int open = css.IndexOf('(', varMatch.Index);
            if (open < 0) {
                continue;
            }

            int close = FindMatchingCssParenthesis(css, open);
            if (close < 0 || close > valueEnd || urlIndex >= close) {
                continue;
            }

            int comma = FindTopLevelComma(css, open + 1, close);
            if (comma >= 0 && urlIndex > comma) {
                return DecodeCssEscapes(varMatch.Groups["name"].Value);
            }
        }

        return null;
    }

    private static bool IsImportantDeclarationValue(string css, int valueStart, int valueEnd) {
        int index = valueEnd - 1;
        while (index >= valueStart && char.IsWhiteSpace(css[index])) {
            index--;
        }

        const string Important = "important";
        if (index - Important.Length + 1 < valueStart) {
            return false;
        }

        string suffix = css.Substring(index - Important.Length + 1, Important.Length);
        if (!string.Equals(suffix, Important, StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        index -= Important.Length;
        while (index >= valueStart && char.IsWhiteSpace(css[index])) {
            index--;
        }

        return index >= valueStart && css[index] == '!';
    }

    private static Dictionary<string, List<CssCustomPropertyDefinition>> CloneCustomPropertyDefinitions(IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> definitions) {
        var clone = new Dictionary<string, List<CssCustomPropertyDefinition>>(StringComparer.Ordinal);
        MergeCustomPropertyDefinitionsInto(clone, definitions);
        return clone;
    }

    private static Dictionary<string, List<CssCustomPropertyDefinition>> MergeCustomPropertyDefinitions(
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> first,
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> second) {
        Dictionary<string, List<CssCustomPropertyDefinition>> merged = CloneCustomPropertyDefinitions(first);
        MergeCustomPropertyDefinitionsInto(merged, second);
        return merged;
    }

    private static void MergeCustomPropertyDefinitionsInto(
        IDictionary<string, List<CssCustomPropertyDefinition>> target,
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> source) {
        foreach (KeyValuePair<string, List<CssCustomPropertyDefinition>> pair in source) {
            if (!target.TryGetValue(pair.Key, out List<CssCustomPropertyDefinition>? values)) {
                values = new List<CssCustomPropertyDefinition>();
                target[pair.Key] = values;
            }

            values.AddRange(pair.Value);
        }
    }

    private static List<SourceRange> GetInactiveCssRuleRanges(string css, HtmlResourcePipelineOptions options) {
        List<SourceRange> ranges = GetInactiveMediaRanges(css, options);
        ranges.AddRange(GetInactiveSupportsRanges(css));
        return ranges;
    }

    private static List<SourceRange> GetInactiveMediaRanges(string css, HtmlResourcePipelineOptions options) {
        var ranges = new List<SourceRange>();
        int index = 0;
        while (index < css.Length) {
            int mediaStart = css.IndexOf("@media", index, StringComparison.OrdinalIgnoreCase);
            if (mediaStart < 0) {
                break;
            }

            if (IsInsideCssString(css, mediaStart) || !HasAtRuleTokenBoundary(css, mediaStart, "@media")) {
                index = mediaStart + 6;
                continue;
            }

            int preludeStart = mediaStart + 6;
            int open = FindNextTopLevelBlockStart(css, preludeStart);
            if (open < 0) {
                break;
            }

            int close = FindMatchingCssBrace(css, open);
            if (close <= open) {
                break;
            }

            string mediaText = css.Substring(preludeStart, open - preludeStart).Trim();
            if (!IsApplicableMedia(mediaText, options)) {
                ranges.Add(new SourceRange(open + 1, close));
                index = close + 1;
            } else {
                index = open + 1;
            }
        }

        return ranges;
    }

    private static List<SourceRange> GetInactiveSupportsRanges(string css) {
        var ranges = new List<SourceRange>();
        int index = 0;
        while (index < css.Length) {
            int supportsStart = css.IndexOf("@supports", index, StringComparison.OrdinalIgnoreCase);
            if (supportsStart < 0) {
                break;
            }

            if (IsInsideCssString(css, supportsStart) || !HasAtRuleTokenBoundary(css, supportsStart, "@supports")) {
                index = supportsStart + 9;
                continue;
            }

            int preludeStart = supportsStart + 9;
            int open = FindNextTopLevelBlockStart(css, preludeStart);
            if (open < 0) {
                break;
            }

            int close = FindMatchingCssBrace(css, open);
            if (close <= open) {
                break;
            }

            string conditionText = css.Substring(preludeStart, open - preludeStart).Trim();
            if (!HtmlComputedStyleEngine.IsApplicableSupports(conditionText)) {
                ranges.Add(new SourceRange(open + 1, close));
                index = close + 1;
            } else {
                index = open + 1;
            }
        }

        return ranges;
    }

    private static int FindNextTopLevelBlockStart(string css, int start) {
        int depth = 0;
        char quote = '\0';
        for (int i = start; i < css.Length; i++) {
            char current = css[i];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(css, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
                continue;
            }

            if (current == '(') {
                depth++;
                continue;
            }

            if (current == ')') {
                depth = Math.Max(0, depth - 1);
                continue;
            }

            if (depth == 0) {
                if (current == '{') {
                    return i;
                }

                if (current == ';') {
                    return -1;
                }
            }
        }

        return -1;
    }

    private static int FindMatchingCssBrace(string css, int open) {
        int depth = 0;
        char quote = '\0';
        for (int i = open; i < css.Length; i++) {
            char current = css[i];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(css, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
                continue;
            }

            if (current == '{') {
                depth++;
                continue;
            }

            if (current == '}') {
                depth--;
                if (depth == 0) {
                    return i;
                }
            }
        }

        return -1;
    }

    private static void AddCustomPropertyDefinition(IDictionary<string, List<CssCustomPropertyDefinition>> definitions, string propertyName, string source, string selector, int declarationStart, bool isImportant, IReadOnlyList<string> aliases, bool isInline, IElement? inlineOwner, string valueText, string? fallbackAlias) {
        if (!definitions.TryGetValue(propertyName, out List<CssCustomPropertyDefinition>? values)) {
            values = new List<CssCustomPropertyDefinition>();
            definitions[propertyName] = values;
        }

        values.Add(new CssCustomPropertyDefinition(source, selector, declarationStart, !string.IsNullOrWhiteSpace(source), isImportant, aliases, isInline, inlineOwner, valueText, fallbackAlias));
    }

    private static void AddUsedCustomPropertyUrls(
        HtmlResourceManifest manifest,
        IElement element,
        string attributeName,
        string css,
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> customPropertyDefinitions,
        IReadOnlyDictionary<IElement, int> inlineSourceOrders,
        IReadOnlyList<SourceRange> inactiveRanges,
        Uri? baseUri,
        HtmlResourcePipelineOptions options,
        IHtmlDocument? document,
        IElement? useElement) {
        if (customPropertyDefinitions.Count == 0) {
            return;
        }

        foreach (Match match in CssVarExpression.Matches(css)) {
            string propertyName = DecodeCssEscapes(match.Groups["name"].Value);
            if (IsInRanges(match.Index, inactiveRanges)
                || !IsCssFunctionNameAt(css, match.Index, "var")
                || IsInsideCssString(css, match.Index)) {
                continue;
            }

            HtmlResourceKind kind = ClassifyCssUrl(css, match.Index);
            if (kind == HtmlResourceKind.Other) {
                continue;
            }

            string useSelector = GetDeclarationSelector(css, match.Index);
            if (!IsCssReferenceForMatchingSelector(document, attributeName, css, match.Index)) {
                continue;
            }

            var addedSources = new HashSet<string>(StringComparer.Ordinal);
            if (document != null && useElement == null && !string.Equals(attributeName, "style", StringComparison.OrdinalIgnoreCase)) {
                IElement[] matchedElements = GetElementsMatchingSelectorList(document, useSelector).ToArray();
                if (matchedElements.Length > 0) {
                    foreach (IElement matchedElement in matchedElements) {
                        Dictionary<string, List<CssCustomPropertyDefinition>> inlineDefinitions = ExtractInlineCustomPropertyDefinitions(matchedElement, inlineSourceOrders, options, includeSelf: true);
                        Dictionary<string, List<CssCustomPropertyDefinition>> mergedDefinitions = inlineDefinitions.Count == 0
                            ? CloneCustomPropertyDefinitions(customPropertyDefinitions)
                            : MergeCustomPropertyDefinitions(customPropertyDefinitions, inlineDefinitions);
                        foreach (CssCustomPropertyDefinition source in ResolveCustomPropertyUrlDefinitions(propertyName, mergedDefinitions, useSelector, document, matchedElement, new HashSet<string>(StringComparer.Ordinal), depth: 0)) {
                            if (!IsFragmentOnlyReference(source.Source) && addedSources.Add(source.Source)) {
                                AddRaw(manifest, kind, element, attributeName + "-var-url", source.Source, baseUri, options);
                            }
                        }
                    }

                    continue;
                }
            }

            foreach (CssCustomPropertyDefinition source in ResolveCustomPropertyUrlDefinitions(propertyName, customPropertyDefinitions, useSelector, document, useElement, new HashSet<string>(StringComparer.Ordinal), depth: 0)) {
                if (!IsFragmentOnlyReference(source.Source) && addedSources.Add(source.Source)) {
                    AddRaw(manifest, kind, element, attributeName + "-var-url", source.Source, baseUri, options);
                }
            }
        }
    }

}
