using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static void SetSvgCssWinner(
        IDictionary<string, SvgCssWinner> winners,
        SvgCssDeclaration declaration,
        SvgCssSpecificity specificity,
        int order) {
        if (winners.TryGetValue(declaration.Name, out SvgCssWinner existing)) {
            if (existing.Important != declaration.Important && !declaration.Important) return;
            int specificityOrder = specificity.CompareTo(existing.Specificity);
            if (existing.Important == declaration.Important
                && (specificityOrder < 0 || specificityOrder == 0 && order < existing.Order)) return;
        }
        winners[declaration.Name] = new SvgCssWinner(declaration.Value, declaration.Important, specificity, order);
    }

    private static bool MatchesSvgSelector(
        XElement element,
        IReadOnlyList<SvgSelectorPart> parts,
        XNamespace svgNamespace,
        ref long remainingWork,
        ref bool workExceeded) {
        if (parts.Count == 0) return false;
        XElement? current = element;
        int index = parts.Count - 1;
        if (!TryConsumeSvgCssMatchWork(ref remainingWork, ref workExceeded)) return false;
        if (!MatchesSvgCompound(current, parts[index].Compound, svgNamespace)) return false;
        while (index > 0) {
            bool directParent = parts[index].DirectParent;
            index--;
            if (directParent) {
                current = current.Parent;
                if (!TryConsumeSvgCssMatchWork(ref remainingWork, ref workExceeded)) return false;
                if (current == null || !MatchesSvgCompound(current, parts[index].Compound, svgNamespace)) return false;
            } else {
                current = current.Parent;
                while (current != null) {
                    if (!TryConsumeSvgCssMatchWork(ref remainingWork, ref workExceeded)) return false;
                    if (MatchesSvgCompound(current, parts[index].Compound, svgNamespace)) break;
                    current = current.Parent;
                }
                if (current == null) return false;
            }
        }
        return true;
    }

    private static bool TryParseSvgSelector(string selector, out IReadOnlyList<SvgSelectorPart> parsed) {
        var parts = new List<SvgSelectorPart>();
        int cursor = 0;
        bool directParent = false;
        while (cursor < selector.Length && char.IsWhiteSpace(selector[cursor])) cursor++;
        if (cursor >= selector.Length || selector[cursor] == '>') {
            parsed = Array.Empty<SvgSelectorPart>();
            return false;
        }
        while (cursor < selector.Length) {
            int start = cursor;
            int brackets = 0;
            char quote = '\0';
            while (cursor < selector.Length) {
                char current = selector[cursor];
                if (quote != '\0') {
                    if (current == quote && selector[cursor - 1] != '\\') quote = '\0';
                } else if (current is '\'' or '"') quote = current;
                else if (current == '[') brackets++;
                else if (current == ']') {
                    if (--brackets < 0) {
                        parsed = Array.Empty<SvgSelectorPart>();
                        return false;
                    }
                } else if (brackets == 0 && (current == '>' || char.IsWhiteSpace(current))) break;
                cursor++;
            }
            if (quote != '\0' || brackets != 0 || cursor == start) {
                parsed = Array.Empty<SvgSelectorPart>();
                return false;
            }
            parts.Add(new SvgSelectorPart(selector.Substring(start, cursor - start), directParent));
            bool hadWhitespace = false;
            while (cursor < selector.Length && char.IsWhiteSpace(selector[cursor])) {
                hadWhitespace = true;
                cursor++;
            }
            if (cursor >= selector.Length) break;
            if (selector[cursor] == '>') {
                directParent = true;
                cursor++;
                while (cursor < selector.Length && char.IsWhiteSpace(selector[cursor])) cursor++;
                if (cursor >= selector.Length || selector[cursor] == '>') {
                    parsed = Array.Empty<SvgSelectorPart>();
                    return false;
                }
            } else if (hadWhitespace) {
                directParent = false;
            } else {
                parsed = Array.Empty<SvgSelectorPart>();
                return false;
            }
        }
        parsed = parts.AsReadOnly();
        return parts.Count > 0;
    }

    private static bool MatchesSvgCompound(XElement element, string compound, XNamespace svgNamespace) {
        if (!IsNativeSvgElement(element, svgNamespace)) return false;
        if (compound.Length == 0 || compound.IndexOf(':') >= 0) return false;
        int index = 0;
        if (compound[0] != '#' && compound[0] != '.' && compound[0] != '[') {
            int start = index;
            while (index < compound.Length && compound[index] != '#' && compound[index] != '.' && compound[index] != '[') index++;
            string type = compound.Substring(start, index - start);
            if (type != "*" && !element.Name.LocalName.Equals(type, StringComparison.OrdinalIgnoreCase)) return false;
        }
        while (index < compound.Length) {
            char marker = compound[index++];
            if (marker == '#') {
                string id = ReadSvgSelectorName(compound, ref index);
                if (!string.Equals(element.Attribute("id")?.Value, id, StringComparison.Ordinal)) return false;
            } else if (marker == '.') {
                string className = ReadSvgSelectorName(compound, ref index);
                string[] classes = (element.Attribute("class")?.Value ?? string.Empty).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
                if (!classes.Contains(className, StringComparer.Ordinal)) return false;
            } else if (marker == '[') {
                int close = compound.IndexOf(']', index);
                if (close < 0) return false;
                string predicate = compound.Substring(index, close - index).Trim();
                int equals = predicate.IndexOf('=');
                string name = (equals < 0 ? predicate : predicate.Substring(0, equals)).Trim();
                XAttribute? attribute = element.Attributes().FirstOrDefault(item =>
                    item.Name.NamespaceName.Length == 0 && item.Name.LocalName.Equals(name, StringComparison.Ordinal));
                if (attribute == null) return false;
                if (equals >= 0) {
                    string expected = predicate.Substring(equals + 1).Trim().Trim('\'', '"');
                    if (!string.Equals(attribute.Value, expected, StringComparison.Ordinal)) return false;
                }
                index = close + 1;
            } else return false;
        }
        return true;
    }

    private static string ReadSvgSelectorName(string text, ref int index) {
        int start = index;
        while (index < text.Length && text[index] != '#' && text[index] != '.' && text[index] != '[') index++;
        return text.Substring(start, index - start);
    }

    private static bool TryCalculateSvgSpecificity(
        string selector,
        out SvgCssSpecificity specificity,
        out IReadOnlyList<SvgSelectorPart> parts) {
        specificity = default;
        parts = Array.Empty<SvgSelectorPart>();
        if (selector.IndexOf('+') >= 0 || selector.IndexOf('~') >= 0 || selector.IndexOf(':') >= 0 ||
            selector.IndexOf('|') >= 0 || selector.IndexOf('\\') >= 0 ||
            selector.IndexOf('(') >= 0 || selector.IndexOf(')') >= 0 ||
            selector.IndexOf("^=", StringComparison.Ordinal) >= 0 || selector.IndexOf("$=", StringComparison.Ordinal) >= 0 ||
            selector.IndexOf("*=", StringComparison.Ordinal) >= 0 || selector.IndexOf("|=", StringComparison.Ordinal) >= 0 ||
            selector.IndexOf("~=", StringComparison.Ordinal) >= 0) return false;
        if (!TryParseSvgSelector(selector, out parts)) return false;
        int ids = 0;
        int classes = 0;
        int types = 0;
        foreach (SvgSelectorPart part in parts) {
            string compound = part.Compound;
            if (!HasSupportedSvgCompoundSelectorSyntax(compound)) return false;
            ids += compound.Count(character => character == '#');
            classes += compound.Count(character => character == '.') + compound.Count(character => character == '[');
            if (compound.Length > 0 && compound[0] != '*' && compound[0] != '#' && compound[0] != '.' && compound[0] != '[') types++;
        }
        specificity = new SvgCssSpecificity(0, ids, classes, types);
        return true;
    }

    private static bool TryConsumeSvgCssMatchWork(ref long remainingWork, ref bool exceeded) {
        if (remainingWork-- > 0L) return true;
        exceeded = true;
        return false;
    }

    private static bool HasSupportedSvgCompoundSelectorSyntax(string compound) {
        if (compound.Length == 0) return false;
        int cursor = 0;
        if (compound[cursor] != '#' && compound[cursor] != '.' && compound[cursor] != '[') {
            int nameStart = cursor;
            while (cursor < compound.Length && compound[cursor] != '#' && compound[cursor] != '.' && compound[cursor] != '[') cursor++;
            string typeName = compound.Substring(nameStart, cursor - nameStart);
            if (typeName != "*" && !IsSupportedSvgSelectorIdentifier(typeName)) return false;
        }
        while (cursor < compound.Length) {
            char marker = compound[cursor++];
            if (marker is '#' or '.') {
                int nameStart = cursor;
                while (cursor < compound.Length && compound[cursor] != '#' && compound[cursor] != '.' && compound[cursor] != '[') cursor++;
                if (!IsSupportedSvgSelectorIdentifier(compound.Substring(nameStart, cursor - nameStart))) return false;
                continue;
            }
            if (marker != '[') return false;
            int close = compound.IndexOf(']', cursor);
            if (close < 0) return false;
            string predicate = compound.Substring(cursor, close - cursor).Trim();
            if (predicate.Length == 0 || predicate.IndexOf('[') >= 0) return false;
            int equals = predicate.IndexOf('=');
            string name = (equals < 0 ? predicate : predicate.Substring(0, equals)).Trim();
            if (!IsSupportedSvgSelectorIdentifier(name)) return false;
            if (equals >= 0) {
                string expected = predicate.Substring(equals + 1).Trim();
                if (expected.Length == 0) return false;
                if (expected[0] is '\'' or '"') {
                    if (expected.Length < 2 || expected[expected.Length - 1] != expected[0]) return false;
                } else if (expected.Any(char.IsWhiteSpace) || expected.IndexOfAny(new[] { '\'', '"' }) >= 0) {
                    return false;
                }
            }
            cursor = close + 1;
        }
        return true;
    }

    private static bool IsSupportedSvgSelectorIdentifier(string name) {
        if (name.Length == 0 || name[0] == '-' || char.IsDigit(name[0])) return false;
        foreach (char character in name) {
            if (!char.IsLetterOrDigit(character) && character != '-' && character != '_') return false;
        }
        return true;
    }
}
