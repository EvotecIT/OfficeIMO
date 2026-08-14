using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

public static partial class HtmlResourcePipeline {
    private static bool SelectorMatchesElementOrAncestor(string selector, IElement? useElement) {
        return GetElementSubstitutionRank(selector, useElement) >= 0;
    }

    private static bool ElementMatchesSelector(IElement element, string selector) {
        string normalized = NormalizeSelectorForQuery(selector, stripPseudoElements: false, stripStatefulPseudoClasses: true);
        if (normalized.Length == 0 || normalized.StartsWith("@", StringComparison.Ordinal)) {
            return false;
        }

        try {
            return element.Matches(normalized);
        } catch {
            return false;
        }
    }

    private static bool IsResolvedVarFallbackUrl(
        string css,
        int urlIndex,
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> customPropertyDefinitions,
        IReadOnlyDictionary<IElement, int> inlineSourceOrders,
        IHtmlDocument? document,
        IElement? useElement,
        IReadOnlyList<SourceRange> inactiveRanges,
        HtmlResourcePipelineOptions options,
        string attributeName) {
        if (customPropertyDefinitions.Count == 0) {
            return false;
        }

        foreach (Match match in CssVarExpression.Matches(css)) {
            string propertyName = DecodeCssEscapes(match.Groups["name"].Value);
            if (IsInRanges(match.Index, inactiveRanges)
                || !IsCssFunctionNameAt(css, match.Index, "var")
                || IsInsideCssString(css, match.Index)
                || !customPropertyDefinitions.TryGetValue(propertyName, out List<CssCustomPropertyDefinition>? sources)) {
                continue;
            }

            int open = css.IndexOf('(', match.Index);
            if (open < 0) {
                continue;
            }

            int close = FindMatchingCssParenthesis(css, open);
            if (close <= open) {
                continue;
            }

            int comma = FindTopLevelComma(css, open + 1, close);
            if (comma < 0 || urlIndex <= comma || urlIndex >= close) {
                continue;
            }

            string useSelector = GetDeclarationSelector(css, match.Index);
            if (document != null && useElement == null && !string.Equals(attributeName, "style", StringComparison.OrdinalIgnoreCase)) {
                IElement[] matchedElements = GetElementsMatchingSelectorList(document, useSelector).ToArray();
                if (matchedElements.Length > 0) {
                    return matchedElements.All(matchedElement => HasResolvedCustomProperty(propertyName, customPropertyDefinitions, inlineSourceOrders, document, matchedElement, options, useSelector));
                }
            }

            return HasResolvedCustomProperty(propertyName, customPropertyDefinitions, document, useElement, new HashSet<string>(StringComparer.Ordinal), depth: 0, useSelector: useSelector);
        }

        return false;
    }

    private static bool HasResolvedCustomProperty(
        string propertyName,
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> customPropertyDefinitions,
        IReadOnlyDictionary<IElement, int> inlineSourceOrders,
        IHtmlDocument document,
        IElement useElement,
        HtmlResourcePipelineOptions options,
        string useSelector) {
        Dictionary<string, List<CssCustomPropertyDefinition>> inlineDefinitions = ExtractInlineCustomPropertyDefinitions(useElement, inlineSourceOrders, options, includeSelf: true);
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> mergedDefinitions = inlineDefinitions.Count == 0
            ? customPropertyDefinitions
            : MergeCustomPropertyDefinitions(customPropertyDefinitions, inlineDefinitions);
        return HasResolvedCustomProperty(propertyName, mergedDefinitions, document, useElement, new HashSet<string>(StringComparer.Ordinal), depth: 0, useSelector: useSelector);
    }

    private static bool HasResolvedCustomProperty(
        string propertyName,
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> customPropertyDefinitions,
        IHtmlDocument? document,
        IElement? useElement,
        ISet<string> visited,
        int depth,
        string useSelector = "") {
        if (depth >= MaxCustomPropertyResolutionDepth
            || !visited.Add(propertyName)
            || !customPropertyDefinitions.TryGetValue(propertyName, out List<CssCustomPropertyDefinition>? sources)) {
            return false;
        }

        bool resolved = false;
        int selectedDeclarationStart = SelectCustomPropertyDeclaration(sources, useSelector, document, useElement);
        if (selectedDeclarationStart >= 0) {
            foreach (CssCustomPropertyDefinition source in sources) {
                if (source.DeclarationStart != selectedDeclarationStart || !CanSubstituteCustomProperty(source, useSelector, document, useElement)) {
                    continue;
                }

                if (source.IsInheritedKeyword) {
                    if (useElement?.ParentElement != null) {
                        visited.Remove(propertyName);
                        resolved = HasResolvedCustomProperty(propertyName, customPropertyDefinitions, document, useElement.ParentElement, visited, depth + 1);
                        visited.Add(propertyName);
                    }
                } else if (source.HasUrl) {
                    resolved = source.FallbackAlias == null
                        || !HasResolvedCustomProperty(source.FallbackAlias, customPropertyDefinitions, document, useElement, visited, depth + 1, useSelector);
                } else if (source.Aliases.Count == 0 && !source.IsCssWideInvalidatingKeyword) {
                    resolved = true;
                } else {
                    foreach (string alias in source.Aliases) {
                        if (HasResolvedCustomProperty(alias, customPropertyDefinitions, document, useElement, visited, depth + 1, useSelector)) {
                            resolved = true;
                            break;
                        }
                    }
                }

                if (resolved) {
                    break;
                }
            }
        }

        visited.Remove(propertyName);
        return resolved;
    }

    private static int FindDeclarationValueEnd(string css, int start) {
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

            if (depth == 0 && (current == ';' || current == '}')) {
                return i;
            }
        }

        return css.Length;
    }

    private static int FindTopLevelComma(string css, int start, int end) {
        int depth = 0;
        char quote = '\0';
        for (int i = start; i < end; i++) {
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
                if (depth > 0) {
                    depth--;
                }

                continue;
            }

            if (depth == 0 && current == ',') {
                return i;
            }
        }

        return -1;
    }

    private static bool IsAncestorSelector(string definitionSelector, string useSelector) {
        if (definitionSelector.Length == 0 || useSelector.Length <= definitionSelector.Length) {
            return false;
        }

        if (!useSelector.StartsWith(definitionSelector, StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        char next = useSelector[definitionSelector.Length];
        return IsCssWhitespace(next) || next == '>';
    }

    private static bool IsSameElementSelectorPrefix(string definitionSelector, string useSelector) {
        if (definitionSelector.Length == 0 || useSelector.Length <= definitionSelector.Length) {
            return false;
        }

        if (!useSelector.StartsWith(definitionSelector, StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        char next = useSelector[definitionSelector.Length];
        return next == '.'
            || next == '#'
            || next == '['
            || next == ':';
    }

    private static string GetDeclarationSelector(string css, int index) {
        int blockStart = css.LastIndexOf('{', Math.Max(0, index - 1));
        if (blockStart < 0) {
            return string.Empty;
        }

        int previousBlockEnd = css.LastIndexOf('}', Math.Max(0, blockStart - 1));
        int previousStatementEnd = css.LastIndexOf(';', Math.Max(0, blockStart - 1));
        int selectorStart = Math.Max(0, Math.Max(previousBlockEnd, previousStatementEnd) + 1);
        string selector = css.Substring(selectorStart, blockStart - selectorStart)
            .Replace(CssCommentMask.ToString(), string.Empty)
            .Trim();
        int groupingStart = selector.LastIndexOf('{');
        return groupingStart >= 0
            ? selector.Substring(groupingStart + 1).Trim()
            : selector;
    }

    private static bool IsCssReferenceForMatchingSelector(IHtmlDocument? document, string attributeName, string css, int index) {
        if (document == null || string.Equals(attributeName, "style", StringComparison.OrdinalIgnoreCase)) {
            return true;
        }

        string selector = GetDeclarationSelector(css, index);
        if (string.IsNullOrWhiteSpace(selector) || selector.TrimStart().StartsWith("@", StringComparison.Ordinal)) {
            return true;
        }

        string[] selectorParts = SplitTopLevelList(selector).ToArray();
        if (selectorParts.Length == 0 || selectorParts.Any(string.IsNullOrWhiteSpace)) return false;
        bool potentiallyActive = false;
        foreach (string selectorPart in selectorParts) {
            string normalized = NormalizeSelectorForQuery(selectorPart, stripStatefulPseudoClasses: true);
            if (normalized.Length == 0) {
                if (IsBarePseudoElementSelector(selectorPart) || IsStatefulPseudoClassOnlySelector(selectorPart)) {
                    potentiallyActive = true;
                    continue;
                }

                return false;
            }

            try {
                if (document.QuerySelector(normalized) != null) {
                    potentiallyActive = true;
                }
            } catch {
                return false;
            }
        }

        return potentiallyActive;
    }

    private static IEnumerable<IElement> GetElementsMatchingSelectorList(IHtmlDocument document, string selector) {
        if (string.IsNullOrWhiteSpace(selector) || selector.TrimStart().StartsWith("@", StringComparison.Ordinal)) {
            yield break;
        }

        string[] selectorParts = SplitTopLevelList(selector).ToArray();
        if (selectorParts.Length == 0 || selectorParts.Any(string.IsNullOrWhiteSpace)) yield break;
        var matchedGroups = new List<IElement[]>();
        foreach (string selectorPart in selectorParts) {
            string normalized = NormalizeSelectorForQuery(selectorPart, stripStatefulPseudoClasses: true);
            if (normalized.Length == 0) {
                if (!IsBarePseudoElementSelector(selectorPart) && !IsStatefulPseudoClassOnlySelector(selectorPart)) yield break;
                continue;
            }

            try {
                matchedGroups.Add(document.QuerySelectorAll(normalized).OfType<IElement>().ToArray());
            } catch {
                yield break;
            }
        }

        var seen = new HashSet<IElement>();
        foreach (IElement[] matches in matchedGroups) {
            foreach (IElement match in matches) {
                if (seen.Add(match)) {
                    yield return match;
                }
            }
        }
    }

    private static bool SelectorRelationshipMatches(IHtmlDocument? document, string definitionSelector, string useSelector) {
        if (document == null || string.IsNullOrWhiteSpace(definitionSelector) || string.IsNullOrWhiteSpace(useSelector)) {
            return false;
        }

        string normalizedDefinition = NormalizeSelectorForQuery(definitionSelector, stripPseudoElements: false);
        string normalizedUse = NormalizeSelectorForQuery(useSelector);
        if (normalizedDefinition.Length == 0 || normalizedUse.Length == 0) {
            return false;
        }

        try {
            if (document.QuerySelector(normalizedDefinition + " " + normalizedUse) != null) {
                return true;
            }

            foreach (IElement useMatch in document.QuerySelectorAll(normalizedUse)) {
                for (IElement? ancestor = useMatch.ParentElement; ancestor != null; ancestor = ancestor.ParentElement) {
                    if (ancestor.Matches(normalizedDefinition)) {
                        return true;
                    }
                }
            }

            return false;
        } catch {
            return false;
        }
    }

    private static bool SelectorSameElementMatches(IHtmlDocument? document, string definitionSelector, string useSelector) {
        if (document == null || string.IsNullOrWhiteSpace(definitionSelector) || string.IsNullOrWhiteSpace(useSelector)) {
            return false;
        }

        string normalizedDefinition = NormalizeSelectorForQuery(definitionSelector, stripPseudoElements: false);
        string normalizedUse = NormalizeSelectorForQuery(useSelector);
        if (normalizedDefinition.Length == 0 || normalizedUse.Length == 0) {
            return false;
        }

        try {
            foreach (IElement useMatch in document.QuerySelectorAll(normalizedUse)) {
                if (useMatch.Matches(normalizedDefinition)) {
                    return true;
                }
            }

            return false;
        } catch {
            return false;
        }
    }

    private static string NormalizeSelectorForQuery(string selector, bool stripPseudoElements = true, bool stripStatefulPseudoClasses = false) {
        string normalized = selector.Trim();
        int pseudoElement = stripPseudoElements ? FindPseudoElementStart(normalized) : -1;
        if (pseudoElement >= 0) {
            normalized = normalized.Substring(0, pseudoElement).TrimEnd();
        }

        if (stripStatefulPseudoClasses) {
            normalized = StripStatefulPseudoClasses(normalized).Trim();
        }

        return normalized;
    }

    private static int FindPseudoElementStart(string selector) {
        int parenthesisDepth = 0;
        int bracketDepth = 0;
        char quote = '\0';
        for (int index = 0; index < selector.Length; index++) {
            char current = selector[index];
            if (current == '\\') {
                index++;
                continue;
            }
            if (quote != '\0') {
                if (current == quote) quote = '\0';
                continue;
            }
            if (current is '\'' or '"') { quote = current; continue; }
            if (current == '[') { bracketDepth++; continue; }
            if (current == ']') { bracketDepth = Math.Max(0, bracketDepth - 1); continue; }
            if (current == '(') { parenthesisDepth++; continue; }
            if (current == ')') { parenthesisDepth = Math.Max(0, parenthesisDepth - 1); continue; }
            if (current != ':' || parenthesisDepth != 0 || bracketDepth != 0) continue;
            if (index + 1 < selector.Length && selector[index + 1] == ':') return index;
            if (IsLegacyPseudoElement(selector, index + 1, "before") ||
                IsLegacyPseudoElement(selector, index + 1, "after")) return index;
        }
        return -1;
    }

    private static bool IsLegacyPseudoElement(string selector, int start, string name) {
        if (start > selector.Length - name.Length ||
            !selector.Substring(start, name.Length).Equals(name, StringComparison.OrdinalIgnoreCase)) return false;
        int end = start + name.Length;
        return end == selector.Length || !(char.IsLetterOrDigit(selector[end]) || selector[end] is '-' or '_');
    }

    private static bool IsBarePseudoElementSelector(string selector) {
        string trimmed = selector.Trim();
        return trimmed.StartsWith("::", StringComparison.Ordinal);
    }

    private static bool IsStatefulPseudoClassOnlySelector(string selector) {
        string stripped = StripStatefulPseudoClasses(selector.Trim()).Trim();
        return stripped.Length == 0;
    }

    private static string StripStatefulPseudoClasses(string selector) {
        var result = new StringBuilder(selector.Length);
        for (int i = 0; i < selector.Length; i++) {
            if (selector[i] == ':' && (i + 1 >= selector.Length || selector[i + 1] != ':') &&
                TryReadPseudoClassName(selector, i + 1, out string pseudoClassName, out int nameEnd)) {
                if (IsStatefulPseudoClass(pseudoClassName)) {
                    i = nameEnd - 1;
                    continue;
                }
                if (nameEnd < selector.Length && selector[nameEnd] == '(' && IsSelectorListPseudoClass(pseudoClassName) &&
                    TryFindMatchingParenthesis(selector, nameEnd, out int closingParenthesis) &&
                    ContainsStatefulPseudoClass(selector, nameEnd + 1, closingParenthesis)) {
                    i = closingParenthesis;
                    continue;
                }
            }

            result.Append(selector[i]);
        }

        return result.ToString();
    }

    private static bool IsSelectorListPseudoClass(string name) =>
        name.Equals("is", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("not", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("where", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("has", StringComparison.OrdinalIgnoreCase);

    private static bool ContainsStatefulPseudoClass(string selector, int start, int end) {
        for (int i = start; i < end; i++) {
            if (selector[i] == ':' && (i + 1 >= end || selector[i + 1] != ':') &&
                TryReadPseudoClassName(selector, i + 1, out string name, out _) && IsStatefulPseudoClass(name)) return true;
        }
        return false;
    }

    private static bool TryFindMatchingParenthesis(string selector, int opening, out int closing) {
        int depth = 0;
        char quote = '\0';
        for (int i = opening; i < selector.Length; i++) {
            char current = selector[i];
            if (quote != '\0') {
                if (current == '\\') i++;
                else if (current == quote) quote = '\0';
                continue;
            }
            if (current is '\'' or '"') quote = current;
            else if (current == '(') depth++;
            else if (current == ')' && --depth == 0) {
                closing = i;
                return true;
            }
        }
        closing = -1;
        return false;
    }

    private static bool TryReadPseudoClassName(string selector, int start, out string name, out int end) {
        int cursor = start;
        while (cursor < selector.Length && (char.IsLetterOrDigit(selector[cursor]) || selector[cursor] == '-')) {
            cursor++;
        }

        if (cursor == start) {
            name = string.Empty;
            end = start;
            return false;
        }

        name = selector.Substring(start, cursor - start);
        end = cursor;
        return true;
    }

    private static bool IsStatefulPseudoClass(string pseudoClassName) {
        switch (pseudoClassName.ToLowerInvariant()) {
            case "active":
            case "focus":
            case "focus-visible":
            case "focus-within":
            case "hover":
            case "target":
            case "visited":
                return true;
            default:
                return false;
        }
    }

}
