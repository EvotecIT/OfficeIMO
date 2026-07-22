using AngleSharp.Dom;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private static SelectorCandidateKey GetSelectorCandidateKey(string selector) {
        if (TryParsePseudoElementSelector(selector, out string hostSelector, out _)) selector = hostSelector;
        int start = FindRightmostCompoundStart(selector);
        string compound = selector.Substring(start).Trim();
        if (compound.Length == 0) return new SelectorCandidateKey(SelectorCandidateKind.Universal, string.Empty);

        string id = FindTopLevelSimpleToken(compound, '#');
        if (id.Length > 0) return new SelectorCandidateKey(SelectorCandidateKind.Id, id);
        string className = FindTopLevelSimpleToken(compound, '.');
        if (className.Length > 0) return new SelectorCandidateKey(SelectorCandidateKind.Class, className);

        int length = ReadSimpleIdentifier(compound, 0);
        if (length > 0 && IsCandidateTokenBoundary(compound, length)) {
            return new SelectorCandidateKey(SelectorCandidateKind.Tag, compound.Substring(0, length));
        }

        return new SelectorCandidateKey(SelectorCandidateKind.Universal, string.Empty);
    }

    private static int FindRightmostCompoundStart(string selector) {
        int squareDepth = 0;
        int parenthesisDepth = 0;
        char quote = '\0';
        int start = 0;
        for (int index = 0; index < selector.Length; index++) {
            char current = selector[index];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(selector, index)) quote = '\0';
                continue;
            }
            if (current == '"' || current == '\'') { quote = current; continue; }
            if (current == '[') { squareDepth++; continue; }
            if (current == ']') { if (squareDepth > 0) squareDepth--; continue; }
            if (current == '(') { parenthesisDepth++; continue; }
            if (current == ')') { if (parenthesisDepth > 0) parenthesisDepth--; continue; }
            if (squareDepth == 0 && parenthesisDepth == 0
                && (char.IsWhiteSpace(current) || current == '>' || current == '+' || current == '~')) {
                start = index + 1;
            }
        }
        return start;
    }

    private static string FindTopLevelSimpleToken(string compound, char marker) {
        int squareDepth = 0;
        int parenthesisDepth = 0;
        char quote = '\0';
        for (int index = 0; index < compound.Length; index++) {
            char current = compound[index];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(compound, index)) quote = '\0';
                continue;
            }
            if (current == '"' || current == '\'') { quote = current; continue; }
            if (current == '[') { squareDepth++; continue; }
            if (current == ']') { if (squareDepth > 0) squareDepth--; continue; }
            if (current == '(') { parenthesisDepth++; continue; }
            if (current == ')') { if (parenthesisDepth > 0) parenthesisDepth--; continue; }
            if (squareDepth != 0 || parenthesisDepth != 0 || current != marker || IsEscaped(compound, index)) continue;
            int tokenStart = index + 1;
            int tokenLength = ReadSimpleIdentifier(compound, tokenStart);
            int tokenEnd = tokenStart + tokenLength;
            if (tokenLength > 0 && IsCandidateTokenBoundary(compound, tokenEnd)) {
                return compound.Substring(tokenStart, tokenLength);
            }
            if (tokenLength > 0) index = tokenEnd - 1;
        }
        return string.Empty;
    }

    private static bool IsCandidateTokenBoundary(string compound, int index) =>
        index >= compound.Length
        || compound[index] == '#'
        || compound[index] == '.'
        || compound[index] == '['
        || compound[index] == ':';

    private static int ReadSimpleIdentifier(string value, int start) {
        int index = start;
        while (index < value.Length && IsIdentifierCharacter(value[index])) index++;
        return index - start;
    }

    private static bool MatchesSelector(IElement element, string selector) {
        try {
            return element.Matches(selector);
        } catch (DomException) {
            return MatchesSimpleSelector(element, selector);
        }
    }

    private static bool MatchesSimpleSelector(IElement element, string selector) {
        if (selector.StartsWith(".", StringComparison.Ordinal)) {
            return element.ClassList.Contains(selector.Substring(1));
        }

        if (selector.StartsWith("#", StringComparison.Ordinal)) {
            return string.Equals(element.Id, selector.Substring(1), StringComparison.Ordinal);
        }

        return string.Equals(element.TagName, selector, StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryParsePseudoElementSelector(string selector, out string hostSelector, out HtmlPseudoElementKind kind) {
        string value = selector.TrimEnd();
        if (TryTrimPseudoElement(value, "::before", out hostSelector)
            || TryTrimPseudoElement(value, ":before", out hostSelector)) {
            kind = HtmlPseudoElementKind.Before;
            return true;
        }

        if (TryTrimPseudoElement(value, "::after", out hostSelector)
            || TryTrimPseudoElement(value, ":after", out hostSelector)) {
            kind = HtmlPseudoElementKind.After;
            return true;
        }

        hostSelector = string.Empty;
        kind = HtmlPseudoElementKind.Before;
        return false;
    }

    private static bool TryTrimPseudoElement(string selector, string suffix, out string hostSelector) {
        if (!selector.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)) {
            hostSelector = string.Empty;
            return false;
        }

        hostSelector = selector.Substring(0, selector.Length - suffix.Length).TrimEnd();
        if (hostSelector.Length == 0) hostSelector = "*";
        return true;
    }


    private static Specificity CalculateSpecificity(string selector) {
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
                i = SkipIdentifier(selector, i + 1);
            } else if (current == '.') {
                classesAttributesAndPseudoClasses++;
                i = SkipIdentifier(selector, i + 1);
            } else if (current == ':') {
                if (i + 1 < selector.Length && selector[i + 1] == ':') {
                    elements++;
                    i = SkipIdentifier(selector, i + 2);
                } else {
                    int nameStart = i + 1;
                    int nameEnd = SkipIdentifier(selector, nameStart);
                    string pseudoName = selector.Substring(nameStart, nameEnd - nameStart + 1);
                    if (nameEnd + 1 < selector.Length && selector[nameEnd + 1] == '(') {
                        int close = FindMatchingParenthesis(selector, nameEnd + 1);
                        if (close > nameEnd + 1) {
                            string argument = selector.Substring(nameEnd + 2, close - nameEnd - 2);
                            if (string.Equals(pseudoName, "where", StringComparison.OrdinalIgnoreCase)) {
                                i = close;
                                continue;
                            }

                            if (string.Equals(pseudoName, "is", StringComparison.OrdinalIgnoreCase)
                                || string.Equals(pseudoName, "not", StringComparison.OrdinalIgnoreCase)
                                || string.Equals(pseudoName, "has", StringComparison.OrdinalIgnoreCase)) {
                                Specificity argumentSpecificity = MaxSpecificity(argument);
                                ids += argumentSpecificity.Ids;
                                classesAttributesAndPseudoClasses += argumentSpecificity.ClassesAttributesAndPseudoClasses;
                                elements += argumentSpecificity.Elements;
                                i = close;
                                continue;
                            }

                            classesAttributesAndPseudoClasses++;
                            i = close;
                            continue;
                        }
                    }

                    if (string.Equals(pseudoName, "before", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(pseudoName, "after", StringComparison.OrdinalIgnoreCase)) {
                        elements++;
                        i = nameEnd;
                        continue;
                    }

                    classesAttributesAndPseudoClasses++;
                    i = nameEnd;
                }
            } else if (IsElementStart(selector, i)) {
                elements++;
                i = SkipIdentifier(selector, i);
            }
        }

        return new Specificity(ids, classesAttributesAndPseudoClasses, elements);
    }

    private static Specificity MaxSpecificity(string selectorList) {
        var max = new Specificity(0, 0, 0);
        foreach (string selector in SplitSelectorList(selectorList)) {
            Specificity specificity = CalculateSpecificity(selector);
            if (specificity.CompareTo(max) > 0) {
                max = specificity;
            }
        }

        return max;
    }

    private static int FindMatchingParenthesis(string text, int openIndex) {
        int depth = 0;
        char quote = '\0';
        for (int i = openIndex; i < text.Length; i++) {
            char current = text[i];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(text, i)) {
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
            } else if (current == ')') {
                depth--;
                if (depth == 0) {
                    return i;
                }
            }
        }

        return -1;
    }

    private static int SkipIdentifier(string selector, int index) {
        int current = index;
        while (current < selector.Length && IsIdentifierCharacter(selector[current])) {
            current++;
        }

        return current - 1;
    }

    private static bool IsElementStart(string selector, int index) {
        char current = selector[index];
        if (!char.IsLetter(current)) {
            return false;
        }

        if (index > 0) {
            char previous = selector[index - 1];
            if (previous == '#' || previous == '.' || previous == '-' || previous == '_' || previous == ':') {
                return false;
            }
        }

        return true;
    }

    private static bool IsIdentifierCharacter(char value) {
        return char.IsLetterOrDigit(value) || value == '-' || value == '_';
    }

}
