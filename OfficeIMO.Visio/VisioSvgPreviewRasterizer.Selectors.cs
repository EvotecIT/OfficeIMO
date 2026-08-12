using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    internal static partial class VisioSvgPreviewRasterizer {
        private static class SvgCssSelectorMatcher {
            private const int MaximumMatchSteps = 256;

            internal static bool MayMatch(XElement element, string selector) {
                return Evaluate(element, selector, out _) != SelectorMatch.NoMatch;
            }

            internal static SelectorMatch Evaluate(XElement element, string selector, out SelectorSpecificity specificity) {
                specificity = default;
                if (!TryTokenize(selector, out List<SelectorPart> parts) ||
                    parts.Count == 0 ||
                    !TryCalculateSpecificity(parts, out specificity)) {
                    return SelectorMatch.Unsupported;
                }
                int remainingSteps = MaximumMatchSteps;
                return MatchesPart(element, parts, parts.Count - 1, ref remainingSteps);
            }

            private static SelectorMatch MatchesPart(XElement element, IReadOnlyList<SelectorPart> parts, int index, ref int remainingSteps) {
                if (remainingSteps-- <= 0) return SelectorMatch.Unsupported;
                SelectorMatch compoundMatch = MatchesCompound(element, parts[index].Compound);
                if (compoundMatch != SelectorMatch.Match) return compoundMatch;
                if (index == 0) return SelectorMatch.Match;

                XElement? parent = element.Parent;
                if (parts[index].Combinator == '>') {
                    return parent == null ? SelectorMatch.NoMatch : MatchesPart(parent, parts, index - 1, ref remainingSteps);
                }

                bool foundUnsupported = false;
                while (parent != null) {
                    SelectorMatch parentMatch = MatchesPart(parent, parts, index - 1, ref remainingSteps);
                    if (parentMatch == SelectorMatch.Match) return SelectorMatch.Match;
                    if (parentMatch == SelectorMatch.Unsupported) foundUnsupported = true;
                    parent = parent.Parent;
                }
                return foundUnsupported ? SelectorMatch.Unsupported : SelectorMatch.NoMatch;
            }

            private static SelectorMatch MatchesCompound(XElement element, string compound) {
                int index = 0;
                if (compound[index] == '*') {
                    index++;
                } else if (IsNameStart(compound[index])) {
                    int start = index++;
                    while (index < compound.Length && IsNameCharacter(compound[index])) index++;
                    if (!string.Equals(
                            element.Name.LocalName,
                            compound.Substring(start, index - start),
                            StringComparison.Ordinal)) {
                        return SelectorMatch.NoMatch;
                    }
                }

                while (index < compound.Length) {
                    char marker = compound[index++];
                    if (marker == '.' || marker == '#') {
                        int start = index;
                        while (index < compound.Length && IsNameCharacter(compound[index])) index++;
                        if (start == index) return SelectorMatch.Unsupported;
                        string value = compound.Substring(start, index - start);
                        if (marker == '#') {
                            if (!string.Equals(element.Attribute("id")?.Value, value, StringComparison.Ordinal)) return SelectorMatch.NoMatch;
                        } else if (!HasClass(element, value)) {
                            return SelectorMatch.NoMatch;
                        }
                    } else if (marker == '[') {
                        int close = compound.IndexOf(']', index);
                        if (close < 0) return SelectorMatch.Unsupported;
                        SelectorMatch attributeMatch = MatchesAttribute(element, compound.Substring(index, close - index));
                        if (attributeMatch != SelectorMatch.Match) return attributeMatch;
                        index = close + 1;
                    } else if (marker == ':') {
                        int start = index;
                        while (index < compound.Length && IsNameCharacter(compound[index])) index++;
                        string pseudo = compound.Substring(start, index - start);
                        if (string.Equals(pseudo, "last-child", StringComparison.OrdinalIgnoreCase)) {
                            if (element.Parent?.Elements().LastOrDefault() != element) return SelectorMatch.NoMatch;
                        } else if (string.Equals(pseudo, "first-child", StringComparison.OrdinalIgnoreCase)) {
                            if (element.Parent?.Elements().FirstOrDefault() != element) return SelectorMatch.NoMatch;
                        } else {
                            return SelectorMatch.Unsupported;
                        }
                    } else {
                        return SelectorMatch.Unsupported;
                    }
                }
                return SelectorMatch.Match;
            }

            private static SelectorMatch MatchesAttribute(XElement element, string expression) {
                string trimmed = expression.Trim();
                int equals = trimmed.IndexOf('=');
                string name = (equals < 0 ? trimmed : trimmed.Substring(0, equals)).Trim();
                if (name.Length == 0 || name.Any(character => !IsNameCharacter(character))) return SelectorMatch.Unsupported;
                // CSS unprefixed attribute selectors match only attributes in no namespace.
                // A local-name scan would incorrectly make [href] match xlink:href.
                XAttribute? attribute = element.Attribute(name);
                if (attribute == null) return SelectorMatch.NoMatch;
                if (equals < 0) return SelectorMatch.Match;
                string rawExpected = trimmed.Substring(equals + 1).Trim();
                if (rawExpected.Length == 0) return SelectorMatch.Unsupported;
                string expected;
                if (rawExpected[0] == '"' || rawExpected[0] == '\'') {
                    char quote = rawExpected[0];
                    int closeQuote = rawExpected.IndexOf(quote, 1);
                    if (closeQuote < 0 || rawExpected.Substring(closeQuote + 1).Trim().Length > 0) {
                        // Attribute-selector value modifiers (for example `i` or `s`) are not
                        // evaluated by this preview matcher. Keep them uncertain so active visual
                        // effects are reported instead of silently treated as unmatched.
                        return SelectorMatch.Unsupported;
                    }
                    expected = rawExpected.Substring(1, closeQuote - 1);
                } else {
                    int whitespace = rawExpected.IndexOfAny(new[] { ' ', '\t', '\r', '\n', '\f' });
                    if (whitespace >= 0) return SelectorMatch.Unsupported;
                    expected = rawExpected;
                }
                return string.Equals(attribute.Value, expected, StringComparison.Ordinal)
                    ? SelectorMatch.Match
                    : SelectorMatch.NoMatch;
            }

            private static bool HasClass(XElement element, string expected) {
                string? value = element.Attribute("class")?.Value;
                return value != null && value
                    .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
                    .Any(item => string.Equals(item, expected, StringComparison.Ordinal));
            }

            private static bool TryTokenize(string selector, out List<SelectorPart> parts) {
                parts = new List<SelectorPart>();
                int index = 0;
                char nextCombinator = '\0';
                while (index < selector.Length) {
                    bool hadWhitespace = false;
                    while (index < selector.Length && char.IsWhiteSpace(selector[index])) {
                        hadWhitespace = true;
                        index++;
                    }
                    if (index >= selector.Length) break;
                    if (selector[index] == '>') {
                        if (parts.Count == 0) return false;
                        nextCombinator = '>';
                        index++;
                        continue;
                    }
                    if (selector[index] == '+' || selector[index] == '~') return false;
                    if (hadWhitespace && parts.Count > 0 && nextCombinator == '\0') nextCombinator = ' ';

                    int start = index;
                    int bracketDepth = 0;
                    int parenthesisDepth = 0;
                    while (index < selector.Length) {
                        char value = selector[index];
                        if (value == '[' && parenthesisDepth == 0) bracketDepth++;
                        if (value == ']' && parenthesisDepth == 0) bracketDepth--;
                        if (value == '(' && bracketDepth == 0) parenthesisDepth++;
                        if (value == ')' && bracketDepth == 0) parenthesisDepth--;
                        if (bracketDepth < 0 || parenthesisDepth < 0) return false;
                        if (bracketDepth == 0 && parenthesisDepth == 0 &&
                            (char.IsWhiteSpace(value) || value == '>')) break;
                        index++;
                    }
                    if (bracketDepth != 0 || parenthesisDepth != 0 || start == index) return false;
                    parts.Add(new SelectorPart(selector.Substring(start, index - start), parts.Count == 0 ? '\0' : nextCombinator == '\0' ? ' ' : nextCombinator));
                    nextCombinator = '\0';
                }
                return nextCombinator == '\0';
            }

            private static bool TryCalculateSpecificity(IReadOnlyList<SelectorPart> parts, out SelectorSpecificity specificity) {
                specificity = default;
                int idCount = 0;
                int classCount = 0;
                int typeCount = 0;
                for (int partIndex = 0; partIndex < parts.Count; partIndex++) {
                    string compound = parts[partIndex].Compound;
                    int index = 0;
                    if (compound[index] == '*') {
                        index++;
                    } else if (IsNameStart(compound[index])) {
                        typeCount++;
                        index++;
                        while (index < compound.Length && IsNameCharacter(compound[index])) index++;
                    }

                    while (index < compound.Length) {
                        char marker = compound[index++];
                        if (marker == '.' || marker == '#') {
                            int start = index;
                            while (index < compound.Length && IsNameCharacter(compound[index])) index++;
                            if (start == index) return false;
                            if (marker == '#') idCount++;
                            else classCount++;
                        } else if (marker == '[') {
                            int close = compound.IndexOf(']', index);
                            if (close < 0) return false;
                            classCount++;
                            index = close + 1;
                        } else if (marker == ':') {
                            int start = index;
                            while (index < compound.Length && IsNameCharacter(compound[index])) index++;
                            if (start == index) return false;
                            classCount++;
                            if (index < compound.Length && compound[index] == '(') {
                                int depth = 1;
                                index++;
                                while (index < compound.Length && depth > 0) {
                                    if (compound[index] == '(') depth++;
                                    else if (compound[index] == ')') depth--;
                                    index++;
                                }
                                if (depth != 0) return false;
                            }
                        } else {
                            return false;
                        }
                    }
                }
                specificity = new SelectorSpecificity(idCount, classCount, typeCount);
                return true;
            }

            private static bool IsNameStart(char value) => char.IsLetter(value) || value == '_' || value == '-';

            private static bool IsNameCharacter(char value) => IsNameStart(value) || char.IsDigit(value);

            internal readonly struct SelectorSpecificity {
                internal SelectorSpecificity(int idCount, int classCount, int typeCount) {
                    IdCount = idCount;
                    ClassCount = classCount;
                    TypeCount = typeCount;
                }

                internal int IdCount { get; }

                internal int ClassCount { get; }

                internal int TypeCount { get; }

                internal int CompareTo(SelectorSpecificity other) {
                    int result = IdCount.CompareTo(other.IdCount);
                    if (result != 0) return result;
                    result = ClassCount.CompareTo(other.ClassCount);
                    return result != 0 ? result : TypeCount.CompareTo(other.TypeCount);
                }
            }

            private readonly struct SelectorPart {
                internal SelectorPart(string compound, char combinator) {
                    Compound = compound;
                    Combinator = combinator;
                }

                internal string Compound { get; }

                internal char Combinator { get; }
            }

            internal enum SelectorMatch {
                NoMatch,
                Match,
                Unsupported
            }
        }
    }
}
