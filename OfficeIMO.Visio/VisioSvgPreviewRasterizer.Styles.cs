using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    internal static partial class VisioSvgPreviewRasterizer {
        private sealed class SvgStyleSheet {
            private const int MaximumSelectorEvaluations = 100000;
            private readonly List<SvgStyleRule> _rules;
            private readonly List<SvgVisualEffectRule> _visualEffectRules;
            private readonly Dictionary<XElement, Dictionary<string, string>> _styleCache = new();
            private readonly Dictionary<XElement, bool> _visualEffectCache = new();
            private int _selectorEvaluations;

            private SvgStyleSheet(List<SvgStyleRule> rules, List<SvgVisualEffectRule> visualEffectRules) {
                _rules = rules;
                _visualEffectRules = visualEffectRules;
            }

            internal static SvgStyleSheet Parse(XElement root) {
                List<SvgStyleRule> rules = new();
                List<SvgVisualEffectRule> visualEffectRules = new();
                int sourceOrder = 0;
                foreach (XElement styleElement in root.Descendants()) {
                    if (!string.Equals(styleElement.Name.LocalName, "style", StringComparison.OrdinalIgnoreCase)) {
                        continue;
                    }

                    ReadRules(styleElement.Value, rules, visualEffectRules, ref sourceOrder);
                }

                return new SvgStyleSheet(rules, visualEffectRules);
            }

            internal bool SelectorBudgetExceeded { get; private set; }

            internal Dictionary<string, string> CreateStyle(XElement element) {
                if (_styleCache.TryGetValue(element, out Dictionary<string, string>? cached)) return cached;
                Dictionary<string, string> style = new(StringComparer.OrdinalIgnoreCase);
                Dictionary<string, (bool Important, int Specificity, int Order)> applied = new(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < _rules.Count; i++) {
                    if (!TryConsumeSelectorEvaluation()) break;
                    SvgStyleRule rule = _rules[i];
                    if (rule.Matches(element)) {
                        MergeRuleDeclarations(style, applied, rule.Declarations, rule.Specificity, rule.Order);
                    }
                }

                MergeDeclarations(style, element.Attribute("style")?.Value);
                _styleCache[element] = style;
                return style;
            }

            internal bool TryGetValue(XElement element, string name, out string? value) {
                Dictionary<string, string> style = CreateStyle(element);
                return style.TryGetValue(name, out value);
            }

            private static void ReadRules(
                string? css,
                List<SvgStyleRule> rules,
                List<SvgVisualEffectRule> visualEffectRules,
                ref int sourceOrder) {
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
                        int ruleOrder = sourceOrder++;
                        string[] selectors = selectorList.Split(',');
                        for (int i = 0; i < selectors.Length; i++) {
                            if ((declarations.ContainsKey("filter") || declarations.ContainsKey("mask")) &&
                                !string.IsNullOrWhiteSpace(selectors[i])) {
                                visualEffectRules.Add(new SvgVisualEffectRule(
                                    selectors[i].Trim(),
                                    new Dictionary<string, string>(declarations, StringComparer.OrdinalIgnoreCase),
                                    ruleOrder));
                            }
                            if (TryCreateRule(selectors[i], declarations, ruleOrder, out SvgStyleRule? rule) && rule != null) {
                                rules.Add(rule);
                            }
                        }
                    }

                    index = close + 1;
                }
            }

            internal bool HasActiveVisualEffect(XElement element) {
                if (_visualEffectCache.TryGetValue(element, out bool cached)) return cached;
                bool active = HasActiveVisualEffect(element, "filter") || HasActiveVisualEffect(element, "mask");
                _visualEffectCache[element] = active;
                return active;
            }

            private bool HasActiveVisualEffect(XElement element, string propertyName) {
                EffectCandidate candidate = default;
                string? presentationValue = element.Attribute(propertyName)?.Value;
                if (TryParseEffectValue(presentationValue, out string? normalizedPresentation, out bool presentationImportant)) {
                    candidate = new EffectCandidate(normalizedPresentation!, presentationImportant, specificity: 0, order: -1);
                }

                EffectCandidate uncertainActiveCandidate = default;
                for (int i = 0; i < _visualEffectRules.Count; i++) {
                    SvgVisualEffectRule rule = _visualEffectRules[i];
                    if (!rule.Declarations.TryGetValue(propertyName, out string? rawValue) ||
                        !TryParseEffectValue(rawValue, out string? value, out bool important)) {
                        continue;
                    }

                    if (!TryConsumeSelectorEvaluation()) break;

                    SvgCssSelectorMatcher.SelectorMatch match = SvgCssSelectorMatcher.Evaluate(
                        element,
                        rule.Selector,
                        out int specificity);
                    if (match == SvgCssSelectorMatcher.SelectorMatch.NoMatch) continue;
                    if (match == SvgCssSelectorMatcher.SelectorMatch.Unsupported) {
                        if (IsActiveEffectValue(element, propertyName, value!)) {
                            var uncertain = new EffectCandidate(value!, important, specificity, rule.Order);
                            if (!uncertainActiveCandidate.HasValue || uncertain.HasHigherPriorityThan(uncertainActiveCandidate)) {
                                uncertainActiveCandidate = uncertain;
                            }
                        }
                        continue;
                    }

                    var next = new EffectCandidate(value!, important, specificity, rule.Order);
                    if (!candidate.HasValue || next.HasHigherPriorityThan(candidate)) candidate = next;
                }

                EffectCandidate inline = default;
                Dictionary<string, string> inlineDeclarations = ParseDeclarations(element.Attribute("style")?.Value);
                if (inlineDeclarations.TryGetValue(propertyName, out string? inlineValue) &&
                    TryParseEffectValue(inlineValue, out string? normalizedInline, out bool inlineImportant)) {
                    inline = new EffectCandidate(normalizedInline!, inlineImportant, specificity: 1000, order: int.MaxValue);
                    if (!candidate.HasValue || inline.HasHigherPriorityThan(candidate)) candidate = inline;
                }

                if (candidate.HasValue && IsActiveEffectValue(element, propertyName, candidate.Value!)) return true;
                return uncertainActiveCandidate.HasValue &&
                       (!candidate.HasValue || uncertainActiveCandidate.HasHigherPriorityThan(candidate));
            }

            private bool TryConsumeSelectorEvaluation() {
                if (_selectorEvaluations >= MaximumSelectorEvaluations) {
                    SelectorBudgetExceeded = true;
                    return false;
                }

                _selectorEvaluations++;
                return true;
            }

            private static bool TryParseEffectValue(string? raw, out string? value, out bool important) {
                value = null;
                important = false;
                if (string.IsNullOrWhiteSpace(raw)) return false;
                string trimmed = raw!.Trim();
                const string importantSuffix = "!important";
                if (trimmed.EndsWith(importantSuffix, StringComparison.OrdinalIgnoreCase)) {
                    important = true;
                    trimmed = trimmed.Substring(0, trimmed.Length - importantSuffix.Length).TrimEnd();
                }
                if (trimmed.Length == 0) return false;
                value = trimmed;
                return true;
            }

            private bool IsActiveEffectValue(XElement element, string propertyName, string value) {
                string normalized = value.Trim();
                if (string.Equals(normalized, "none", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(normalized, "initial", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(normalized, "unset", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(normalized, "revert", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(normalized, "revert-layer", StringComparison.OrdinalIgnoreCase)) {
                    return false;
                }
                if (string.Equals(normalized, "inherit", StringComparison.OrdinalIgnoreCase)) {
                    return element.Parent != null && HasActiveVisualEffect(element.Parent, propertyName);
                }
                return true;
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
                        if (style.TryGetValue(name, out string? existing) &&
                            IsImportantDeclaration(existing) &&
                            !IsImportantDeclaration(value)) {
                            continue;
                        }
                        style[name] = value;
                    }
                }
            }

            private static bool IsImportantDeclaration(string value) =>
                value.TrimEnd().EndsWith("!important", StringComparison.OrdinalIgnoreCase);

            private static void MergeRuleDeclarations(
                Dictionary<string, string> style,
                Dictionary<string, (bool Important, int Specificity, int Order)> applied,
                Dictionary<string, string> declarations,
                int specificity,
                int order) {
                foreach (KeyValuePair<string, string> declaration in declarations) {
                    bool important = IsImportantDeclaration(declaration.Value);
                    if (!applied.TryGetValue(declaration.Key, out (bool Important, int Specificity, int Order) previous) ||
                        important && !previous.Important ||
                        important == previous.Important &&
                        (specificity > previous.Specificity ||
                         specificity == previous.Specificity && order >= previous.Order)) {
                        style[declaration.Key] = declaration.Value;
                        applied[declaration.Key] = (important, specificity, order);
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

            private sealed class SvgVisualEffectRule {
                internal SvgVisualEffectRule(string selector, Dictionary<string, string> declarations, int order) {
                    Selector = selector;
                    Declarations = declarations;
                    Order = order;
                }

                internal string Selector { get; }

                internal Dictionary<string, string> Declarations { get; }

                internal int Order { get; }
            }

            private readonly struct EffectCandidate {
                internal EffectCandidate(string value, bool important, int specificity, int order) {
                    Value = value;
                    Important = important;
                    Specificity = specificity;
                    Order = order;
                    HasValue = true;
                }

                internal bool HasValue { get; }

                internal bool Important { get; }

                internal int Order { get; }

                internal int Specificity { get; }

                internal string? Value { get; }

                internal bool HasHigherPriorityThan(EffectCandidate other) =>
                    Important != other.Important
                        ? Important
                        : Specificity != other.Specificity
                            ? Specificity > other.Specificity
                            : Order >= other.Order;
            }
        }
    }
}
