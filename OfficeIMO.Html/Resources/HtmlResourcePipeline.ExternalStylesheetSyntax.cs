using System.Runtime.CompilerServices;

namespace OfficeIMO.Html;

public static partial class HtmlResourcePipeline {
    // The publication keeps decoded resource strings alive. Cache only lexical work here;
    // media, resource policy and budgets still belong to the current render.
    internal sealed class ExternalStylesheetCache {
        private readonly ConditionalWeakTable<string, ExternalStylesheetSyntax> _sheets = new();
        private readonly ConditionalWeakTable<string, IReadOnlyList<(int Start, int Length, string Source)>> _rebaseUrls = new();
        private readonly Dictionary<string, string> _materialized = new(StringComparer.Ordinal);
        private int _materializedCharacters;

        internal HtmlExternalStylesheetAnalysis Analyze(string css, Uri uri, HtmlResourcePipelineOptions options, bool includeInactiveResources) =>
            AnalyzeExternalStylesheet(css, uri, options, includeInactiveResources, Get(css, options.Limits));

        internal IReadOnlyList<(int Start, int Length, string Source)> GetRebaseUrls(string css) =>
            _rebaseUrls.GetValue(ReuseMaterializedText(css), static source => ReadRebaseUrls(source).ToArray());

        internal string ReuseMaterializedText(string css) {
            if (_materialized.TryGetValue(css, out string? existing)) return existing;
            if (css.Length <= 16 * 1024 * 1024 - _materializedCharacters && _materialized.Count < 4096) {
                _materialized.Add(css, css);
                _materializedCharacters += css.Length;
            }
            return css;
        }

        internal IReadOnlyList<HtmlCssFontFaceDefinition> ExtractFonts(string css, HtmlResourcePipelineOptions options) {
            var definitions = new List<HtmlCssFontFaceDefinition>();
            Get(css, options.Limits).FontFaces.AddApplicable(options, includeInactiveResources: false, definitions);
            return definitions;
        }

        private ExternalStylesheetSyntax Get(string css, HtmlConversionLimits limits) =>
            _sheets.GetValue(css, source => new ExternalStylesheetSyntax(source, limits));
    }

    private sealed class ExternalStylesheetSyntax {
        private readonly (SourceRange Range, string Condition)[] _media;
        private readonly SourceRange[] _unsupported;

        internal ExternalStylesheetSyntax(string source, HtmlConversionLimits limits) {
            Text = StripCssCommentsOutsideStrings(source);
            Imports = ExtractCssImports(Text).ToArray();
            Images = ReadStylesheetImages(Text, Imports.Select(item => new SourceRange(item.Start, item.End)).ToList()).ToArray();
            FontFaces = FontFaceSheet.Parse(Text, limits);
            _media = ReadConditionalRanges(Text, "media").ToArray();
            _unsupported = ReadConditionalRanges(Text, "supports")
                .Where(item => !HtmlComputedStyleEngine.IsApplicableSupports(item.Condition)).Select(item => item.Range).ToArray();
        }

        internal string Text { get; }
        internal CssImportReference[] Imports { get; }
        internal StylesheetImageReference[] Images { get; }
        internal FontFaceSheet FontFaces { get; }

        internal List<SourceRange> GetInactiveRanges(HtmlResourcePipelineOptions options) {
            var ranges = new List<SourceRange>(_unsupported);
            int lastEnd = 0;
            foreach (var condition in _media) {
                if (condition.Range.Start < lastEnd || IsApplicableMedia(condition.Condition, options)) continue;
                ranges.Add(condition.Range);
                lastEnd = condition.Range.End;
            }
            return ranges;
        }

        private static IEnumerable<(SourceRange Range, string Condition)> ReadConditionalRanges(string css, string rule) {
            int index = 0;
            while (index < css.Length && TryFindNextAtRule(css, index, rule, out _, out int prelude)) {
                int open = FindNextTopLevelBlockStart(css, prelude);
                if (open < 0) yield break;
                int close = FindMatchingCssBrace(css, open);
                if (close <= open) yield break;
                yield return (new SourceRange(open + 1, close), css.Substring(prelude, open - prelude).Trim());
                index = open + 1;
            }
        }
    }

    private readonly struct StylesheetImageReference {
        internal StylesheetImageReference(string source, string attribute, int start, bool customProperty, bool pageRule) {
            Source = source;
            Attribute = attribute;
            Start = start;
            CustomProperty = customProperty;
            PageRule = pageRule;
        }

        internal string Source { get; }
        internal string Attribute { get; }
        internal int Start { get; }
        internal bool CustomProperty { get; }
        internal bool PageRule { get; }
    }
}
