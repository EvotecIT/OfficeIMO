using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

public static partial class HtmlResourcePipeline {
    private static IEnumerable<CssCustomPropertyDefinition> ResolveCustomPropertyUrlDefinitions(
        string propertyName,
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> customPropertyDefinitions,
        string useSelector,
        IHtmlDocument? document,
        IElement? useElement,
        ISet<string> visited,
        int depth) {
        if (depth >= MaxCustomPropertyResolutionDepth
            || !visited.Add(propertyName)
            || !customPropertyDefinitions.TryGetValue(propertyName, out List<CssCustomPropertyDefinition>? sources)) {
            yield break;
        }

        int selectedDeclarationStart = SelectCustomPropertyDeclaration(sources, useSelector, document, useElement);
        if (selectedDeclarationStart < 0) {
            visited.Remove(propertyName);
            yield break;
        }

        foreach (CssCustomPropertyDefinition source in sources) {
            if (source.DeclarationStart != selectedDeclarationStart || !CanSubstituteCustomProperty(source, useSelector, document, useElement)) {
                continue;
            }

            if (source.IsInheritedKeyword) {
                foreach (CssCustomPropertyDefinition inheritedSource in ResolveInheritedCustomPropertyUrlDefinitions(propertyName, customPropertyDefinitions, document, useElement, visited, depth)) {
                    yield return inheritedSource;
                }

                continue;
            }

            if (source.HasUrl) {
                if (source.FallbackAlias == null || !HasResolvedCustomProperty(source.FallbackAlias, customPropertyDefinitions, document, useElement, visited, depth + 1)) {
                    yield return source;
                }
            }

            foreach (string alias in source.Aliases) {
                foreach (CssCustomPropertyDefinition aliasSource in ResolveCustomPropertyUrlDefinitions(alias, customPropertyDefinitions, useSelector, document, useElement, visited, depth + 1)) {
                    yield return aliasSource;
                }
            }
        }

        visited.Remove(propertyName);
    }

    private static IEnumerable<CssCustomPropertyDefinition> ResolveInheritedCustomPropertyUrlDefinitions(
        string propertyName,
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> customPropertyDefinitions,
        IHtmlDocument? document,
        IElement? useElement,
        ISet<string> visited,
        int depth) {
        if (useElement?.ParentElement == null) {
            yield break;
        }

        visited.Remove(propertyName);
        foreach (CssCustomPropertyDefinition inheritedSource in ResolveCustomPropertyUrlDefinitions(propertyName, customPropertyDefinitions, string.Empty, document, useElement.ParentElement, visited, depth + 1)) {
            yield return inheritedSource;
        }

        visited.Add(propertyName);
    }

    private static bool CanSubstituteCustomProperty(CssCustomPropertyDefinition source, string useSelector, IHtmlDocument? document = null, IElement? useElement = null) {
        string definitionSelector = source.Selector;
        if (string.IsNullOrWhiteSpace(definitionSelector)) {
            if (source.IsInline && useElement != null) {
                return GetInlineOwnerDistance(source, useElement) != int.MaxValue;
            }

            return string.IsNullOrWhiteSpace(useSelector);
        }

        if (string.Equals(definitionSelector, useSelector, StringComparison.OrdinalIgnoreCase)) {
            return true;
        }

        foreach (string definitionPart in SplitTopLevelList(definitionSelector)) {
            string normalizedDefinition = definitionPart.Trim();
            if (SelectorMatchesElementOrAncestor(normalizedDefinition, useElement)) {
                return true;
            }

            if (string.Equals(normalizedDefinition, ":root", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalizedDefinition, "html", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalizedDefinition, "body", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            if (useElement != null) {
                continue;
            }

            foreach (string usePart in SplitTopLevelList(useSelector)) {
                string normalizedUse = usePart.Trim();
                if (IsAncestorSelector(normalizedDefinition, normalizedUse)
                    || SelectorSameElementMatches(document, normalizedDefinition, normalizedUse)
                    || IsSameElementSelectorPrefix(normalizedDefinition, normalizedUse)
                    || SelectorRelationshipMatches(document, normalizedDefinition, normalizedUse)) {
                    return true;
                }
            }
        }

        return false;
    }

    private static int SelectCustomPropertyDeclaration(IEnumerable<CssCustomPropertyDefinition> sources, string useSelector, IHtmlDocument? document = null, IElement? useElement = null) {
        int selectedDeclarationStart = -1;
        int selectedRank = -1;
        int selectedSpecificity = -1;
        int selectedDistance = int.MaxValue;
        bool selectedImportant = false;
        foreach (CssCustomPropertyDefinition source in sources) {
            int rank = GetSubstitutionRank(source, useSelector, document, useElement);
            if (rank < 0) {
                continue;
            }

            int distance = GetElementSubstitutionDistance(source, useElement);
            int specificity = GetMatchingSelectorSpecificity(source.Selector, useSelector, document, useElement);
            bool sameElementCascade = rank >= 3 && selectedRank >= 3;
            if ((sameElementCascade && source.IsImportant != selectedImportant && source.IsImportant)
                || (!(sameElementCascade && source.IsImportant != selectedImportant) && rank > selectedRank)
                || (rank == selectedRank
                    && (distance < selectedDistance
                        || (distance == selectedDistance
                            && ((!selectedImportant && source.IsImportant)
                                || (source.IsImportant == selectedImportant
                                    && (specificity > selectedSpecificity
                                        || (specificity == selectedSpecificity && source.DeclarationStart >= selectedDeclarationStart)))))))) {
                selectedImportant = source.IsImportant;
                selectedRank = rank;
                selectedSpecificity = specificity;
                selectedDistance = distance;
                selectedDeclarationStart = source.DeclarationStart;
            }
        }

        return selectedRank >= 0 ? selectedDeclarationStart : -1;
    }

    private static int GetSubstitutionRank(CssCustomPropertyDefinition source, string useSelector, IHtmlDocument? document = null, IElement? useElement = null) {
        string definitionSelector = source.Selector;
        if (string.IsNullOrWhiteSpace(definitionSelector)) {
            if (source.IsInline && useElement != null) {
                int inlineDistance = GetInlineOwnerDistance(source, useElement);
                if (inlineDistance == int.MaxValue) {
                    return -1;
                }

                return inlineDistance == 0 ? 4 : 2;
            }

            return string.IsNullOrWhiteSpace(useSelector) ? 3 : -1;
        }

        int best = -1;
        foreach (string definitionPart in SplitTopLevelList(definitionSelector)) {
            string normalizedDefinition = definitionPart.Trim();
            best = Math.Max(best, GetElementSubstitutionRank(normalizedDefinition, useElement));
            if (string.Equals(normalizedDefinition, ":root", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalizedDefinition, "html", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalizedDefinition, "body", StringComparison.OrdinalIgnoreCase)) {
                best = Math.Max(best, 1);
            }

            if (useElement != null) {
                continue;
            }

            foreach (string usePart in SplitTopLevelList(useSelector)) {
                string normalizedUse = usePart.Trim();
                if (string.Equals(normalizedDefinition, normalizedUse, StringComparison.OrdinalIgnoreCase)) {
                    best = Math.Max(best, 3);
                } else if (SelectorSameElementMatches(document, normalizedDefinition, normalizedUse)) {
                    best = Math.Max(best, 3);
                } else if (IsAncestorSelector(normalizedDefinition, normalizedUse)
                    || IsSameElementSelectorPrefix(normalizedDefinition, normalizedUse)
                    || SelectorRelationshipMatches(document, normalizedDefinition, normalizedUse)) {
                    best = Math.Max(best, 2);
                }
            }
        }

        return best;
    }

    private static int GetElementSubstitutionRank(string definitionSelector, IElement? useElement) {
        if (useElement == null || string.IsNullOrWhiteSpace(definitionSelector)) {
            return -1;
        }

        if (ElementMatchesSelector(useElement, definitionSelector)) {
            return 3;
        }

        for (IElement? ancestor = useElement.ParentElement; ancestor != null; ancestor = ancestor.ParentElement) {
            if (ElementMatchesSelector(ancestor, definitionSelector)) {
                return 2;
            }
        }

        return -1;
    }

    private static int GetElementSubstitutionDistance(CssCustomPropertyDefinition source, IElement? useElement) {
        if (source.IsInline) {
            return GetInlineOwnerDistance(source, useElement);
        }

        string definitionSelector = source.Selector;
        if (useElement == null || string.IsNullOrWhiteSpace(definitionSelector)) {
            return int.MaxValue;
        }

        int best = int.MaxValue;
        foreach (string definitionPart in SplitTopLevelList(definitionSelector)) {
            string normalizedDefinition = definitionPart.Trim();
            if (ElementMatchesSelector(useElement, normalizedDefinition)) {
                best = Math.Min(best, 0);
                continue;
            }

            int distance = 1;
            for (IElement? ancestor = useElement.ParentElement; ancestor != null; ancestor = ancestor.ParentElement, distance++) {
                if (ElementMatchesSelector(ancestor, normalizedDefinition)) {
                    best = Math.Min(best, distance);
                    break;
                }
            }
        }

        return best;
    }

    private static int GetInlineOwnerDistance(CssCustomPropertyDefinition source, IElement? useElement) {
        if (!source.IsInline || source.InlineOwner == null || useElement == null) {
            return int.MaxValue;
        }

        int distance = 0;
        for (IElement? current = useElement; current != null; current = current.ParentElement, distance++) {
            if (ReferenceEquals(current, source.InlineOwner)) {
                return distance;
            }
        }

        return int.MaxValue;
    }

    private static int GetMatchingSelectorSpecificity(string definitionSelector, string useSelector, IHtmlDocument? document, IElement? useElement) {
        int best = -1;
        foreach (string definitionPart in SplitTopLevelList(definitionSelector)) {
            string normalizedDefinition = definitionPart.Trim();
            if (normalizedDefinition.Length == 0) {
                continue;
            }

            bool matches = SelectorMatchesElementOrAncestor(normalizedDefinition, useElement)
                || string.Equals(normalizedDefinition, ":root", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalizedDefinition, "html", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalizedDefinition, "body", StringComparison.OrdinalIgnoreCase);
            if (!matches) {
                if (useElement != null) {
                    continue;
                }

                foreach (string usePart in SplitTopLevelList(useSelector)) {
                    string normalizedUse = usePart.Trim();
                    if (string.Equals(normalizedDefinition, normalizedUse, StringComparison.OrdinalIgnoreCase)
                        || SelectorSameElementMatches(document, normalizedDefinition, normalizedUse)
                        || IsAncestorSelector(normalizedDefinition, normalizedUse)
                        || IsSameElementSelectorPrefix(normalizedDefinition, normalizedUse)
                        || SelectorRelationshipMatches(document, normalizedDefinition, normalizedUse)) {
                        matches = true;
                        break;
                    }
                }
            }

            if (matches) {
                best = Math.Max(best, CalculateSelectorSpecificity(normalizedDefinition));
            }
        }

        return best;
    }

    private static int CalculateSelectorSpecificity(string selector) {
        int ids = 0;
        int classesAttributesAndPseudoClasses = 0;
        int elements = 0;
        bool inAttribute = false;
        for (int i = 0; i < selector.Length; i++) {
            char current = selector[i];
            if (current == '[') {
                inAttribute = true;
                classesAttributesAndPseudoClasses++;
                continue;
            }

            if (current == ']') {
                inAttribute = false;
                continue;
            }

            if (inAttribute) {
                continue;
            }

            if (current == '#') {
                ids++;
                i = SkipCssIdentifier(selector, i + 1);
            } else if (current == '.') {
                classesAttributesAndPseudoClasses++;
                i = SkipCssIdentifier(selector, i + 1);
            } else if (current == ':') {
                if (i + 1 < selector.Length && selector[i + 1] == ':') {
                    elements++;
                    i = SkipCssIdentifier(selector, i + 2);
                } else {
                    if (TryReadPseudoClassName(selector, i + 1, out string pseudoClassName, out int nameEnd)) {
                        if (nameEnd < selector.Length && selector[nameEnd] == '(') {
                            int close = FindMatchingCssParenthesis(selector, nameEnd);
                            if (close > nameEnd) {
                                string argument = selector.Substring(nameEnd + 1, close - nameEnd - 1);
                                if (string.Equals(pseudoClassName, "where", StringComparison.OrdinalIgnoreCase)) {
                                    i = close;
                                    continue;
                                }

                                if (string.Equals(pseudoClassName, "is", StringComparison.OrdinalIgnoreCase)
                                    || string.Equals(pseudoClassName, "not", StringComparison.OrdinalIgnoreCase)
                                    || string.Equals(pseudoClassName, "has", StringComparison.OrdinalIgnoreCase)) {
                                    int argumentSpecificity = MaxSelectorSpecificity(argument);
                                    ids += argumentSpecificity / 10000;
                                    classesAttributesAndPseudoClasses += (argumentSpecificity % 10000) / 100;
                                    elements += argumentSpecificity % 100;
                                    i = close;
                                    continue;
                                }

                                classesAttributesAndPseudoClasses++;
                                i = close;
                                continue;
                            }
                        }

                        classesAttributesAndPseudoClasses++;
                        i = nameEnd - 1;
                        continue;
                    }

                    classesAttributesAndPseudoClasses++;
                    i = SkipCssIdentifier(selector, i + 1);
                }
            } else if (IsSelectorElementStart(selector, i)) {
                elements++;
                i = SkipCssIdentifier(selector, i);
            }
        }

        return (ids * 10000) + (classesAttributesAndPseudoClasses * 100) + elements;
    }

    private static int MaxSelectorSpecificity(string selectorList) {
        int max = 0;
        foreach (string selector in SplitTopLevelList(selectorList)) {
            int specificity = CalculateSelectorSpecificity(selector);
            if (specificity > max) {
                max = specificity;
            }
        }

        return max;
    }

    private static int SkipCssIdentifier(string selector, int start) {
        int cursor = start;
        while (cursor < selector.Length && (IsCssIdentifierCharacter(selector[cursor]) || selector[cursor] == '\\')) {
            cursor++;
        }

        return Math.Max(start, cursor) - 1;
    }

    private static bool IsSelectorElementStart(string selector, int index) {
        char current = selector[index];
        if (!char.IsLetter(current) && current != '*') {
            return false;
        }

        if (current == '*') {
            return false;
        }

        if (index > 0) {
            char previous = selector[index - 1];
            if (previous == '#'
                || previous == '.'
                || previous == ':'
                || previous == '-'
                || previous == '_'
                || char.IsLetterOrDigit(previous)) {
                return false;
            }
        }

        return true;
    }

}
