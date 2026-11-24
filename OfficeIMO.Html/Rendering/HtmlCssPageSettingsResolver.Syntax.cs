using System.Runtime.CompilerServices;
using OfficeIMO.Html.Css;

namespace OfficeIMO.Html;

internal static partial class HtmlCssPageSettingsResolver {
    internal sealed class SyntaxCache {
        private readonly ConditionalWeakTable<string, IReadOnlyList<RawPageRule>> _sheets = new();

        internal IReadOnlyList<RawPageRule> Get(string css, CancellationToken token) =>
            _sheets.GetValue(css, source => ReadRawPageRules(source, 0, source.Length, token).ToArray());
    }

    internal sealed class RawPageRule {
        private readonly string? _source;
        private readonly int _start;
        private readonly int _end;
        private RawPageRule[]? _children;
        private PageRuleBody? _body;

        internal RawPageRule(string name, string prelude, string? source = null, int start = 0, int end = 0) {
            Name = name;
            Prelude = prelude;
            _source = source;
            _start = start;
            _end = end;
        }

        internal string Name { get; }
        internal string Prelude { get; }
        internal bool HasBlock => _source != null;

        internal PageRuleBody ReadPageBody(HtmlCssProcessingBudget fontBudget, CancellationToken token) =>
            _body ??= new PageRuleBody(Prelude, _source!.Substring(_start, _end - _start), fontBudget, token);

        // A conditional block may apply to only some chapters. Parse it when first used.
        internal IEnumerable<RawPageRule> ReadChildren(bool cache, CancellationToken token) =>
            cache ? _children ??= ReadRawPageRules(_source!, _start, _end, token).ToArray()
                : ReadRawPageRules(_source!, _start, _end, token);
    }

    internal sealed class PageRuleBody {
        internal PageRuleBody(string selector, string body, HtmlCssProcessingBudget fontBudget, CancellationToken token) {
            HtmlCssStyleBlock syntax = ParseContextStyleBlock(body, fontBudget, token);
            Declarations = new PageDeclarations(syntax, isMarginBox: false, token);
            var diagnostics = new HtmlDiagnosticReport();
            MarginBoxes = ExtractMarginBoxes(body, selector, diagnostics, fontBudget, token);
            StyleDeclarations = ReadContextStyleDeclarations(syntax, token);
            Diagnostics = diagnostics.Diagnostics;
        }

        internal PageDeclarations Declarations { get; }
        internal IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginRule> MarginBoxes { get; }
        internal IReadOnlyList<HtmlGeneratedStyleDeclaration> StyleDeclarations { get; }
        internal IReadOnlyList<HtmlDiagnostic> Diagnostics { get; }
    }

    internal sealed class PageDeclarations {
        private static readonly HashSet<string> PageProperties = new(StringComparer.OrdinalIgnoreCase) {
            "size", "margin", "margin-top", "margin-right", "margin-bottom", "margin-left",
            "background", "background-color", "bleed", "marks"
        };

        private readonly Dictionary<string, List<HtmlCssPageDeclaration>> _properties = new(StringComparer.OrdinalIgnoreCase);

        internal PageDeclarations(HtmlCssStyleBlock syntax, bool isMarginBox, CancellationToken token) {
            string body = HtmlCssTokenizer.StripComments(syntax.Source, syntax.Tokens, token);
            int start = 0;
            int depth = 0;
            char quote = '\0';
            int order = 0;
            for (int index = 0; index <= body.Length; index++) {
                if ((index & 1023) == 0) token.ThrowIfCancellationRequested();
                char current = index < body.Length ? body[index] : ';';
                if (quote != '\0') {
                    if (current == quote && (index == 0 || body[index - 1] != '\\')) quote = '\0';
                    continue;
                }

                if (current == '\'' || current == '"') quote = current;
                else if (current == '(' || current == '{') depth++;
                else if ((current == ')' || current == '}') && depth > 0) depth--;
                else if (current == ';' && depth == 0) {
                    string declaration = body.Substring(start, index - start).Trim();
                    int separator = declaration.IndexOf(':');
                    if (separator > 0) {
                        string name = declaration.Substring(0, separator).Trim();
                        if (isMarginBox ? name.Equals("content", StringComparison.OrdinalIgnoreCase) : PageProperties.Contains(name)) {
                            string value = declaration.Substring(separator + 1).Trim();
                            bool important = TryStripImportant(ref value);
                            if (!_properties.TryGetValue(name, out var candidates))
                                _properties.Add(name, candidates = new List<HtmlCssPageDeclaration>());
                            candidates.Add(new HtmlCssPageDeclaration(value, important, order));
                        }
                    }
                    start = index + 1;
                    order++;
                }
            }
        }

        internal HtmlCssPageDeclaration Find(string name, Func<string, bool>? isValid = null) {
            var resolved = new HtmlCssPageDeclaration(string.Empty, false, -1);
            if (_properties.TryGetValue(name, out var candidates)) {
                foreach (HtmlCssPageDeclaration candidate in candidates) {
                    if ((isValid == null || isValid(candidate.Value)) && (candidate.IsImportant || !resolved.IsImportant))
                        resolved = candidate;
                }
            }
            return resolved;
        }
    }

    private static IEnumerable<RawPageRule> ReadRawPageRules(string css, int start, int end, CancellationToken token) {
        int cursor = start;
        while (cursor < end) {
            token.ThrowIfCancellationRequested();
            if (IsCommentStart(css, cursor)) {
                cursor = SkipComment(css, cursor + 2, end);
                continue;
            }
            char current = css[cursor];
            if (current == '\'' || current == '"') {
                cursor = SkipQuoted(css, cursor + 1, end, current);
                continue;
            }
            if (current == '{') {
                int close = FindMatchingBrace(css, cursor, token);
                cursor = close < 0 ? end : close + 1;
                continue;
            }
            if (current != '@') {
                cursor++;
                continue;
            }

            int nameStart = cursor + 1;
            int nameEnd = nameStart;
            while (nameEnd < end && (char.IsLetter(css[nameEnd]) || css[nameEnd] == '-')) nameEnd++;
            string name = css.Substring(nameStart, nameEnd - nameStart).ToLowerInvariant();
            int boundary = FindRuleBoundary(css, nameEnd, end, token);
            if (boundary < 0) yield break;
            string prelude = css.Substring(nameEnd, boundary - nameEnd);
            if (css[boundary] == ';') {
                if (name == "layer") yield return new RawPageRule(name, prelude);
                cursor = boundary + 1;
                continue;
            }

            int closeBrace = FindMatchingBrace(css, boundary, token);
            if (closeBrace < 0 || closeBrace >= end) yield break;
            if (name is "media" or "supports" or "layer" or "page")
                yield return new RawPageRule(name, prelude.Trim(), css, boundary + 1, closeBrace);
            cursor = closeBrace + 1;
        }
    }

    private static void ApplyRawPageRules(IEnumerable<RawPageRule> rules, HtmlRenderOptions options, HtmlDiagnosticReport diagnostics,
        HtmlCssPageRuleSet pageRules, CascadeLayerRegistry layers, string? layerPath, CascadeLayerOrder? layerOrder,
        HtmlCssProcessingBudget fontBudget, bool cache, CancellationToken token) {
        foreach (RawPageRule rule in rules) {
            token.ThrowIfCancellationRequested();
            switch (rule.Name) {
                case "layer" when !rule.HasBlock:
                    layers.RegisterStatement(rule.Prelude, layerPath);
                    break;
                case "layer":
                    (string path, CascadeLayerOrder order) = layers.RegisterBlock(rule.Prelude, layerPath);
                    ApplyRawPageRules(rule.ReadChildren(cache, token), options, diagnostics, pageRules, layers, path, order, fontBudget, cache, token);
                    break;
                case "media" when IsApplicablePageMedia(rule.Prelude, options):
                case "supports" when HtmlComputedStyleEngine.IsApplicableSupports(rule.Prelude):
                    ApplyRawPageRules(rule.ReadChildren(cache, token), options, diagnostics, pageRules, layers, layerPath, layerOrder, fontBudget, cache, token);
                    break;
                case "page":
                    ApplyRawPageRule(rule, options, diagnostics, pageRules, layerOrder, fontBudget, token);
                    break;
            }
        }
    }
}
