using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private const int MaximumSvgCssRules = 4096;
    private const int MaximumSvgCssDeclarations = 32768;

    private static void ApplySvgStylesheets(XElement root, ref int unsupported) {
        var rules = new List<SvgCssRule>();
        int declarations = 0;
        foreach (XElement style in root.DescendantsAndSelf().Where(element =>
                     element.Name.LocalName.Equals("style", StringComparison.OrdinalIgnoreCase))) {
            ParseSvgCssRules(RemoveSvgCssComments(style.Value), rules, ref declarations, ref unsupported);
            if (rules.Count >= MaximumSvgCssRules || declarations >= MaximumSvgCssDeclarations) break;
        }
        if (rules.Count == 0) return;

        foreach (XElement element in root.DescendantsAndSelf()) {
            if (element.Name.LocalName.Equals("style", StringComparison.OrdinalIgnoreCase)) continue;
            var winners = new Dictionary<string, SvgCssWinner>(StringComparer.OrdinalIgnoreCase);
            foreach (SvgCssRule rule in rules) {
                if (!MatchesSvgSelector(element, rule.Selector)) continue;
                foreach (SvgCssDeclaration declaration in rule.Declarations) {
                    SetSvgCssWinner(winners, declaration, rule.Specificity, rule.Order);
                }
            }
            int inlineOrder = rules.Count + 1;
            foreach (SvgCssDeclaration declaration in ParseSvgCssDeclarations(element.Attribute("style")?.Value, ref declarations)) {
                SetSvgCssWinner(winners, declaration, int.MaxValue, inlineOrder++);
            }
            if (winners.Count == 0) continue;

            IReadOnlyDictionary<string, string> customProperties = ResolveSvgCustomProperties(element, winners);
            var css = new StringBuilder();
            foreach (KeyValuePair<string, SvgCssWinner> pair in winners.OrderBy(item => item.Value.Order)) {
                if (pair.Key.StartsWith("--", StringComparison.Ordinal)) continue;
                if (!TryResolveSvgCssVariables(pair.Value.Value, customProperties, 0, out string value)) {
                    unsupported++;
                    continue;
                }
                if (css.Length > 0) css.Append(';');
                css.Append(pair.Key).Append(':').Append(value);
            }
            foreach (KeyValuePair<string, string> custom in customProperties) {
                if (css.Length > 0) css.Append(';');
                css.Append(custom.Key).Append(':').Append(custom.Value);
            }
            element.SetAttributeValue("style", css.ToString());
        }
    }

    private static IReadOnlyDictionary<string, string> ResolveSvgCustomProperties(
        XElement element,
        IReadOnlyDictionary<string, SvgCssWinner> winners) {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        XElement? parent = element.Parent;
        if (parent != null) {
            foreach (SvgCssDeclaration declaration in ParseSvgCssDeclarations(parent.Attribute("style")?.Value)) {
                if (declaration.Name.StartsWith("--", StringComparison.Ordinal)) result[declaration.Name] = declaration.Value;
            }
        }
        foreach (KeyValuePair<string, SvgCssWinner> winner in winners) {
            if (winner.Key.StartsWith("--", StringComparison.Ordinal)) result[winner.Key] = winner.Value.Value;
        }
        return result;
    }

    private static void ParseSvgCssRules(
        string css,
        ICollection<SvgCssRule> rules,
        ref int declarationCount,
        ref int unsupported) {
        int cursor = 0;
        while (cursor < css.Length && rules.Count < MaximumSvgCssRules && declarationCount < MaximumSvgCssDeclarations) {
            int open = FindSvgCssCharacter(css, '{', cursor);
            if (open < 0) break;
            int close = FindSvgCssBlockEnd(css, open + 1);
            if (close < 0) {
                unsupported++;
                break;
            }
            string selectorText = css.Substring(cursor, open - cursor).Trim();
            string body = css.Substring(open + 1, close - open - 1);
            if (!selectorText.StartsWith("@", StringComparison.Ordinal)) {
                IReadOnlyList<SvgCssDeclaration> declarations = ParseSvgCssDeclarations(body, ref declarationCount);
                foreach (string selector in SplitSvgCssTopLevel(selectorText, ',')) {
                    string normalized = selector.Trim();
                    if (normalized.Length == 0 || declarations.Count == 0) continue;
                    if (!TryCalculateSvgSpecificity(normalized, out int specificity)) {
                        unsupported++;
                        continue;
                    }
                    rules.Add(new SvgCssRule(normalized, declarations, specificity, rules.Count));
                    if (rules.Count >= MaximumSvgCssRules) break;
                }
            }
            cursor = close + 1;
        }
    }

    private static IReadOnlyList<SvgCssDeclaration> ParseSvgCssDeclarations(string? text) {
        int ignored = 0;
        return ParseSvgCssDeclarations(text, ref ignored);
    }

    private static IReadOnlyList<SvgCssDeclaration> ParseSvgCssDeclarations(string? text, ref int declarationCount) {
        var result = new List<SvgCssDeclaration>();
        if (string.IsNullOrWhiteSpace(text)) return result;
        foreach (string raw in SplitSvgCssTopLevel(text!, ';')) {
            if (declarationCount >= MaximumSvgCssDeclarations) break;
            int colon = raw.IndexOf(':');
            if (colon <= 0) continue;
            string name = raw.Substring(0, colon).Trim();
            string value = raw.Substring(colon + 1).Trim();
            if (name.Length == 0 || value.Length == 0) continue;
            bool important = TryStripImportant(value, out value);
            result.Add(new SvgCssDeclaration(name, value, important));
            declarationCount++;
        }
        return result;
    }

    private static void SetSvgCssWinner(
        IDictionary<string, SvgCssWinner> winners,
        SvgCssDeclaration declaration,
        int specificity,
        int order) {
        if (winners.TryGetValue(declaration.Name, out SvgCssWinner existing)) {
            if (existing.Important != declaration.Important && !declaration.Important) return;
            if (existing.Important == declaration.Important
                && (specificity < existing.Specificity || specificity == existing.Specificity && order < existing.Order)) return;
        }
        winners[declaration.Name] = new SvgCssWinner(declaration.Value, declaration.Important, specificity, order);
    }

    private static bool MatchesSvgSelector(XElement element, string selector) {
        IReadOnlyList<SvgSelectorPart> parts = ParseSvgSelector(selector);
        if (parts.Count == 0) return false;
        XElement? current = element;
        for (int index = parts.Count - 1; index >= 0; index--) {
            if (current == null || !MatchesSvgCompound(current, parts[index].Compound)) return false;
            if (index == 0) return true;
            if (parts[index].DirectParent) {
                current = current.Parent;
                continue;
            }
            XElement? ancestor = current.Parent;
            while (ancestor != null && !MatchesSvgCompound(ancestor, parts[index - 1].Compound)) ancestor = ancestor.Parent;
            if (ancestor == null) return false;
            current = ancestor;
            index--;
        }
        return true;
    }

    private static IReadOnlyList<SvgSelectorPart> ParseSvgSelector(string selector) {
        var parts = new List<SvgSelectorPart>();
        var token = new StringBuilder();
        bool direct = false;
        int brackets = 0;
        for (int index = 0; index <= selector.Length; index++) {
            char current = index < selector.Length ? selector[index] : ' ';
            if (current == '[') brackets++;
            if (current == ']') brackets--;
            if (brackets == 0 && (current == '>' || char.IsWhiteSpace(current))) {
                if (token.Length > 0) {
                    parts.Add(new SvgSelectorPart(token.ToString(), direct));
                    token.Clear();
                    direct = false;
                }
                if (current == '>') direct = true;
                continue;
            }
            token.Append(current);
        }
        return parts;
    }

    private static bool MatchesSvgCompound(XElement element, string compound) {
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
                XAttribute? attribute = element.Attribute(name);
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

    private static bool TryCalculateSvgSpecificity(string selector, out int specificity) {
        specificity = 0;
        if (selector.IndexOf('+') >= 0 || selector.IndexOf('~') >= 0 || selector.IndexOf(':') >= 0) return false;
        foreach (SvgSelectorPart part in ParseSvgSelector(selector)) {
            string compound = part.Compound;
            specificity += compound.Count(character => character == '#') * 10000;
            specificity += (compound.Count(character => character == '.') + compound.Count(character => character == '[')) * 100;
            if (compound.Length > 0 && compound[0] != '*' && compound[0] != '#' && compound[0] != '.' && compound[0] != '[') specificity++;
        }
        return true;
    }

    private static bool TryResolveSvgCssVariables(
        string value,
        IReadOnlyDictionary<string, string> customProperties,
        int depth,
        out string resolved) {
        resolved = value;
        if (depth > 16) return false;
        int start = resolved.IndexOf("var(", StringComparison.OrdinalIgnoreCase);
        while (start >= 0) {
            int close = FindSvgCssBlockEnd(resolved, start + 4, '(', ')');
            if (close < 0) return false;
            string arguments = resolved.Substring(start + 4, close - start - 4);
            IReadOnlyList<string> parts = SplitSvgCssTopLevel(arguments, ',');
            string name = parts[0].Trim();
            string replacement;
            if (!customProperties.TryGetValue(name, out replacement!)) {
                if (parts.Count < 2) return false;
                replacement = string.Join(",", parts.Skip(1)).Trim();
            }
            if (!TryResolveSvgCssVariables(replacement, customProperties, depth + 1, out replacement)) return false;
            resolved = resolved.Substring(0, start) + replacement + resolved.Substring(close + 1);
            start = resolved.IndexOf("var(", StringComparison.OrdinalIgnoreCase);
        }
        return true;
    }

    private static IReadOnlyList<string> SplitSvgCssTopLevel(string text, char separator) {
        var result = new List<string>();
        int start = 0;
        int depth = 0;
        char quote = '\0';
        for (int index = 0; index < text.Length; index++) {
            char current = text[index];
            if (quote != '\0') {
                if (current == quote && (index == 0 || text[index - 1] != '\\')) quote = '\0';
                continue;
            }
            if (current is '\'' or '"') quote = current;
            else if (current is '(' or '[') depth++;
            else if (current is ')' or ']') depth--;
            else if (current == separator && depth == 0) {
                result.Add(text.Substring(start, index - start));
                start = index + 1;
            }
        }
        result.Add(text.Substring(start));
        return result;
    }

    private static int FindSvgCssCharacter(string text, char target, int start) {
        char quote = '\0';
        for (int index = start; index < text.Length; index++) {
            char current = text[index];
            if (quote != '\0') {
                if (current == quote && text[index - 1] != '\\') quote = '\0';
            } else if (current is '\'' or '"') quote = current;
            else if (current == target) return index;
        }
        return -1;
    }

    private static int FindSvgCssBlockEnd(string text, int start, char open = '{', char close = '}') {
        int depth = 1;
        char quote = '\0';
        for (int index = start; index < text.Length; index++) {
            char current = text[index];
            if (quote != '\0') {
                if (current == quote && text[index - 1] != '\\') quote = '\0';
            } else if (current is '\'' or '"') quote = current;
            else if (current == open) depth++;
            else if (current == close && --depth == 0) return index;
        }
        return -1;
    }

    private static string RemoveSvgCssComments(string css) {
        var result = new StringBuilder(css.Length);
        for (int index = 0; index < css.Length; index++) {
            if (index + 1 < css.Length && css[index] == '/' && css[index + 1] == '*') {
                int close = css.IndexOf("*/", index + 2, StringComparison.Ordinal);
                if (close < 0) break;
                index = close + 1;
            } else result.Append(css[index]);
        }
        return result.ToString();
    }

    private readonly struct SvgCssDeclaration {
        internal SvgCssDeclaration(string name, string value, bool important) { Name = name; Value = value; Important = important; }
        internal string Name { get; }
        internal string Value { get; }
        internal bool Important { get; }
    }

    private readonly struct SvgCssWinner {
        internal SvgCssWinner(string value, bool important, int specificity, int order) { Value = value; Important = important; Specificity = specificity; Order = order; }
        internal string Value { get; }
        internal bool Important { get; }
        internal int Specificity { get; }
        internal int Order { get; }
    }

    private readonly struct SvgCssRule {
        internal SvgCssRule(string selector, IReadOnlyList<SvgCssDeclaration> declarations, int specificity, int order) { Selector = selector; Declarations = declarations; Specificity = specificity; Order = order; }
        internal string Selector { get; }
        internal IReadOnlyList<SvgCssDeclaration> Declarations { get; }
        internal int Specificity { get; }
        internal int Order { get; }
    }

    private readonly struct SvgSelectorPart {
        internal SvgSelectorPart(string compound, bool directParent) { Compound = compound; DirectParent = directParent; }
        internal string Compound { get; }
        internal bool DirectParent { get; }
    }
}
