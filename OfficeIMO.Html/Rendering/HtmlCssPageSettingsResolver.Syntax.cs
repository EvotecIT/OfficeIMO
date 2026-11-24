using System.Runtime.CompilerServices;

namespace OfficeIMO.Html;

internal static partial class HtmlCssPageSettingsResolver {
    internal sealed class SyntaxCache {
        private readonly ConditionalWeakTable<string, IReadOnlyList<RawPageRule>> _sheets = new();

        internal IReadOnlyList<RawPageRule> Get(string css, CancellationToken token) =>
            _sheets.GetValue(css, source => FreezeRules(ReadRawPageRules(source, 0, source.Length, token)));
    }

    private static IReadOnlyList<RawPageRule> FreezeRules(IEnumerable<RawPageRule> rules) =>
        rules.Select(rule => new RawPageRule(rule.Name, rule.Prelude, rule.Body,
            rule.Children == null ? null : FreezeRules(rule.Children))).ToArray();

    internal sealed class RawPageRule {
        internal RawPageRule(string name, string prelude, string? body, IEnumerable<RawPageRule>? children) {
            Name = name;
            Prelude = prelude;
            Body = body;
            Children = children;
        }

        internal string Name { get; }
        internal string Prelude { get; }
        internal string? Body { get; }
        internal IEnumerable<RawPageRule>? Children { get; }
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
                int close = FindMatchingBrace(css, cursor);
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
            int boundary = FindRuleBoundary(css, nameEnd, end);
            if (boundary < 0) yield break;
            string prelude = css.Substring(nameEnd, boundary - nameEnd);
            if (css[boundary] == ';') {
                if (name == "layer") yield return new RawPageRule(name, prelude, null, null);
                cursor = boundary + 1;
                continue;
            }

            int closeBrace = FindMatchingBrace(css, boundary);
            if (closeBrace < 0 || closeBrace >= end) yield break;
            if (name is "media" or "supports" or "layer")
                yield return new RawPageRule(name, prelude.Trim(), null, ReadRawPageRules(css, boundary + 1, closeBrace, token));
            else if (name == "page")
                yield return new RawPageRule(name, prelude.Trim(), css.Substring(boundary + 1, closeBrace - boundary - 1), null);
            cursor = closeBrace + 1;
        }
    }

    private static void ApplyRawPageRules(IEnumerable<RawPageRule> rules, HtmlRenderOptions options, HtmlDiagnosticReport diagnostics,
        HtmlCssPageRuleSet pageRules, CascadeLayerRegistry layers, string? layerPath, CascadeLayerOrder? layerOrder,
        HtmlCssProcessingBudget fontBudget, CancellationToken token) {
        foreach (RawPageRule rule in rules) {
            token.ThrowIfCancellationRequested();
            switch (rule.Name) {
                case "layer" when rule.Children == null:
                    layers.RegisterStatement(rule.Prelude, layerPath);
                    break;
                case "layer":
                    (string path, CascadeLayerOrder order) = layers.RegisterBlock(rule.Prelude, layerPath);
                    ApplyRawPageRules(rule.Children!, options, diagnostics, pageRules, layers, path, order, fontBudget, token);
                    break;
                case "media" when IsApplicablePageMedia(rule.Prelude, options):
                case "supports" when HtmlComputedStyleEngine.IsApplicableSupports(rule.Prelude):
                    ApplyRawPageRules(rule.Children!, options, diagnostics, pageRules, layers, layerPath, layerOrder, fontBudget, token);
                    break;
                case "page":
                    ApplyRawPageRule(rule.Prelude, rule.Body!, options, diagnostics, pageRules, layerOrder, fontBudget, token);
                    break;
            }
        }
    }
}
