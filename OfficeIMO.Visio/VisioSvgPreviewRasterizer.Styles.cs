using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    internal static partial class VisioSvgPreviewRasterizer {
        private sealed class SvgStyleSheet {
            private readonly List<SvgStyleRule> _rules;

            private SvgStyleSheet(List<SvgStyleRule> rules) {
                _rules = rules;
            }

            internal static SvgStyleSheet Parse(XElement root) {
                List<SvgStyleRule> rules = new();
                foreach (XElement styleElement in root.Descendants()) {
                    if (!string.Equals(styleElement.Name.LocalName, "style", StringComparison.OrdinalIgnoreCase)) {
                        continue;
                    }

                    ReadRules(styleElement.Value, rules);
                }

                return new SvgStyleSheet(rules);
            }

            internal Dictionary<string, string> CreateStyle(XElement element) {
                Dictionary<string, string> style = new(StringComparer.OrdinalIgnoreCase);
                Dictionary<string, (int Specificity, int Order)> applied = new(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < _rules.Count; i++) {
                    SvgStyleRule rule = _rules[i];
                    if (rule.Matches(element)) {
                        MergeRuleDeclarations(style, applied, rule.Declarations, rule.Specificity, rule.Order);
                    }
                }

                MergeDeclarations(style, element.Attribute("style")?.Value);
                return style;
            }

            internal bool TryGetValue(XElement element, string name, out string? value) {
                Dictionary<string, string> style = CreateStyle(element);
                return style.TryGetValue(name, out value);
            }

            private static void ReadRules(string? css, List<SvgStyleRule> rules) {
                if (string.IsNullOrWhiteSpace(css)) {
                    return;
                }

                string normalized = RemoveComments(css!);
                int index = 0;
                while (index < normalized.Length) {
                    int open = normalized.IndexOf('{', index);
                    if (open < 0) {
                        break;
                    }

                    int close = normalized.IndexOf('}', open + 1);
                    if (close < 0) {
                        break;
                    }

                    string selectorList = normalized.Substring(index, open - index);
                    Dictionary<string, string> declarations = ParseDeclarations(normalized.Substring(open + 1, close - open - 1));
                    if (declarations.Count > 0) {
                        string[] selectors = selectorList.Split(',');
                        for (int i = 0; i < selectors.Length; i++) {
                            if (TryCreateRule(selectors[i], declarations, rules.Count, out SvgStyleRule? rule) && rule != null) {
                                rules.Add(rule);
                            }
                        }
                    }

                    index = close + 1;
                }
            }

            private static Dictionary<string, string> ParseDeclarations(string? raw) {
                Dictionary<string, string> style = new(StringComparer.OrdinalIgnoreCase);
                MergeDeclarations(style, raw);
                return style;
            }

            private static void MergeDeclarations(Dictionary<string, string> style, string? raw) {
                if (string.IsNullOrWhiteSpace(raw)) {
                    return;
                }

                string[] declarations = raw!.Split(';');
                for (int i = 0; i < declarations.Length; i++) {
                    int separator = declarations[i].IndexOf(':');
                    if (separator <= 0) {
                        continue;
                    }

                    string name = declarations[i].Substring(0, separator).Trim();
                    string value = declarations[i].Substring(separator + 1).Trim();
                    if (name.Length > 0 && value.Length > 0) {
                        style[name] = value;
                    }
                }
            }

            private static void MergeRuleDeclarations(
                Dictionary<string, string> style,
                Dictionary<string, (int Specificity, int Order)> applied,
                Dictionary<string, string> declarations,
                int specificity,
                int order) {
                foreach (KeyValuePair<string, string> declaration in declarations) {
                    if (!applied.TryGetValue(declaration.Key, out (int Specificity, int Order) previous) ||
                        specificity > previous.Specificity ||
                        specificity == previous.Specificity && order >= previous.Order) {
                        style[declaration.Key] = declaration.Value;
                        applied[declaration.Key] = (specificity, order);
                    }
                }
            }

            private static string RemoveComments(string css) {
                int start = css.IndexOf("/*", StringComparison.Ordinal);
                if (start < 0) {
                    return css;
                }

                StringBuilder builder = new(css.Length);
                int index = 0;
                while (index < css.Length) {
                    start = css.IndexOf("/*", index, StringComparison.Ordinal);
                    if (start < 0) {
                        builder.Append(css, index, css.Length - index);
                        break;
                    }

                    builder.Append(css, index, start - index);
                    int end = css.IndexOf("*/", start + 2, StringComparison.Ordinal);
                    if (end < 0) {
                        break;
                    }

                    index = end + 2;
                }

                return builder.ToString();
            }

            private static bool TryCreateRule(string selector, Dictionary<string, string> declarations, int order, out SvgStyleRule? rule) {
                rule = null;
                string trimmed = selector.Trim();
                if (trimmed.Length == 0 || ContainsUnsupportedSelectorSyntax(trimmed)) {
                    return false;
                }

                string? elementName = null;
                string? id = null;
                List<string> classes = new();
                int index = 0;
                if (index < trimmed.Length && IsNameStartCharacter(trimmed[index])) {
                    int start = index++;
                    while (index < trimmed.Length && IsNameCharacter(trimmed[index])) {
                        index++;
                    }

                    elementName = trimmed.Substring(start, index - start);
                }

                while (index < trimmed.Length) {
                    char marker = trimmed[index];
                    if (marker != '.' && marker != '#') {
                        return false;
                    }

                    index++;
                    int start = index;
                    while (index < trimmed.Length && IsNameCharacter(trimmed[index])) {
                        index++;
                    }

                    if (index == start) {
                        return false;
                    }

                    string value = trimmed.Substring(start, index - start);
                    if (marker == '#') {
                        id = value;
                    } else {
                        classes.Add(value);
                    }
                }

                if (elementName == null && id == null && classes.Count == 0) {
                    return false;
                }

                rule = new SvgStyleRule(elementName, id, classes, new Dictionary<string, string>(declarations, StringComparer.OrdinalIgnoreCase), order);
                return true;
            }

            private static bool ContainsUnsupportedSelectorSyntax(string selector) {
                for (int i = 0; i < selector.Length; i++) {
                    char c = selector[i];
                    if (char.IsWhiteSpace(c) || c == '>' || c == '+' || c == '~' || c == ':' || c == '*' || c == '[') {
                        return true;
                    }
                }

                return false;
            }

            private static bool IsNameStartCharacter(char value) =>
                char.IsLetter(value) || value == '_' || value == '-';

            private static bool IsNameCharacter(char value) =>
                char.IsLetterOrDigit(value) || value == '-' || value == '_';

            private sealed class SvgStyleRule {
                internal SvgStyleRule(string? elementName, string? id, IReadOnlyList<string> classes, Dictionary<string, string> declarations, int order) {
                    ElementName = elementName;
                    Id = id;
                    Classes = classes;
                    Declarations = declarations;
                    Order = order;
                    Specificity = (id == null ? 0 : 100) + (classes.Count * 10) + (elementName == null ? 0 : 1);
                }

                private string? ElementName { get; }

                private string? Id { get; }

                private IReadOnlyList<string> Classes { get; }

                internal Dictionary<string, string> Declarations { get; }

                internal int Order { get; }

                internal int Specificity { get; }

                internal bool Matches(XElement element) {
                    if (ElementName != null && !string.Equals(element.Name.LocalName, ElementName, StringComparison.OrdinalIgnoreCase)) {
                        return false;
                    }

                    if (Id != null && !string.Equals(element.Attribute("id")?.Value, Id, StringComparison.Ordinal)) {
                        return false;
                    }

                    if (Classes.Count == 0) {
                        return true;
                    }

                    string? classAttribute = element.Attribute("class")?.Value;
                    if (string.IsNullOrWhiteSpace(classAttribute)) {
                        return false;
                    }

                    string[] elementClasses = classAttribute!.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
                    for (int i = 0; i < Classes.Count; i++) {
                        bool found = false;
                        for (int j = 0; j < elementClasses.Length; j++) {
                            if (string.Equals(elementClasses[j], Classes[i], StringComparison.OrdinalIgnoreCase)) {
                                found = true;
                                break;
                            }
                        }

                        if (!found) {
                            return false;
                        }
                    }

                    return true;
                }
            }
        }
    }
}
