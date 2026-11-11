using AngleSharp.Css.Dom;
using AngleSharp.Css.Parser;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

internal sealed class HtmlCssFontFaceDefinition {
    internal HtmlCssFontFaceDefinition(string familyName, string source, string weight, string stretch, string style, string unicodeRange) {
        FamilyName = CleanFamilyName(familyName);
        Source = source ?? string.Empty;
        Weight = weight ?? string.Empty;
        Stretch = stretch ?? string.Empty;
        Style = style ?? string.Empty;
        UnicodeRange = unicodeRange ?? string.Empty;
    }

    internal string FamilyName { get; }
    internal string Source { get; }
    internal string Weight { get; }
    internal string Stretch { get; }
    internal string Style { get; }
    internal string UnicodeRange { get; }

    internal HtmlCssFontFaceDefinition WithRawDescriptors(IReadOnlyDictionary<string, string> declarations) =>
        new HtmlCssFontFaceDefinition(
            declarations.TryGetValue("font-family", out string? family) ? family : FamilyName,
            declarations.TryGetValue("src", out string? source) ? source : Source,
            declarations.TryGetValue("font-weight", out string? weight) ? weight : Weight,
            declarations.TryGetValue("font-stretch", out string? stretch) ? stretch : Stretch,
            declarations.TryGetValue("font-style", out string? style) ? style : Style,
            declarations.TryGetValue("unicode-range", out string? range) ? range : UnicodeRange);

    private static string CleanFamilyName(string value) {
        string family = (value ?? string.Empty).Trim();
        while (family.Length >= 2
               && ((family[0] == '"' && family[family.Length - 1] == '"')
                   || (family[0] == '\'' && family[family.Length - 1] == '\''))) {
            family = family.Substring(1, family.Length - 2).Trim();
        }

        return family;
    }
}

public static partial class HtmlResourcePipeline {
    internal static IReadOnlyList<HtmlCssFontFaceDefinition> ExtractFontFaces(string css, HtmlResourcePipelineOptions options, bool includeInactiveResources = false) {
        var definitions = new List<HtmlCssFontFaceDefinition>();
        if (string.IsNullOrWhiteSpace(css)) {
            return definitions.AsReadOnly();
        }

        FontFaceSheet sheet = FontFaceSheets.GetOrParse(css, options.Limits);
        sheet.AddApplicable(options, includeInactiveResources, definitions);
        return definitions.AsReadOnly();
    }

    /// <summary>
    /// The @font-face rules of one stylesheet text, independent of the media environment. A sheet
    /// that embeds its fonts is megabytes of text, parsed again by every conversion of the same
    /// document, so large sheets are remembered by their exact text.
    /// </summary>
    private sealed class FontFaceSheet {
        private readonly int _maximumNestingDepth;
        private readonly IReadOnlyList<(HtmlCssFontFaceDefinition Definition, FontFaceCondition[] Conditions)> _faces;

        private FontFaceSheet(int maximumNestingDepth, IReadOnlyList<(HtmlCssFontFaceDefinition, FontFaceCondition[])> faces) {
            _maximumNestingDepth = maximumNestingDepth;
            _faces = faces;
        }

        internal static FontFaceSheet Parse(string css, HtmlConversionLimits? limits) {
            // The validation scan also yields the rule blocks that raw descriptor recovery needs.
            IReadOnlyDictionary<int, int> closures = HtmlCssRuleBlockScanner.Scan(
                css, new HtmlCssProcessingBudget(limits), out int maximumDepth);
            var faces = new List<(HtmlCssFontFaceDefinition, FontFaceCondition[])>();
            // The style cascade parses every stylesheet already; a full parse here only pays off when the
            // text can contain an @font-face rule.
            if (!MayContainFontFaceRule(css)) {
                return new FontFaceSheet(maximumDepth, faces);
            }

            var parser = new CssParser();
            ICssStyleSheet stylesheet = parser.ParseStyleSheet(css);
            IReadOnlyList<IReadOnlyDictionary<string, string>> rawRules = ExtractRawFontFaceDescriptors(css, closures);
            var rawDescriptors = new Dictionary<ICssFontFaceRule, IReadOnlyDictionary<string, string>>();
            int rawRuleIndex = 0;
            foreach (ICssRule rule in stylesheet.Rules) {
                MapRawFontFaceDescriptors(rule, rawRules, rawDescriptors, ref rawRuleIndex);
            }
            var conditions = new List<FontFaceCondition>();
            foreach (ICssRule rule in stylesheet.Rules) {
                CollectFontFaces(rule, faces, conditions, rawDescriptors);
            }
            return new FontFaceSheet(maximumDepth, faces);
        }

        /// <summary>Replays the scan's limit check, then adds the faces whose media and supports conditions apply.</summary>
        internal void AddApplicable(HtmlResourcePipelineOptions options, bool includeInactiveResources,
            ICollection<HtmlCssFontFaceDefinition> definitions) {
            var budget = new HtmlCssProcessingBudget(options.Limits);
            for (int depth = 1; depth <= _maximumNestingDepth; depth++) budget.RecordNestingDepth(depth);
            foreach ((HtmlCssFontFaceDefinition definition, FontFaceCondition[] conditions) in _faces) {
                if (includeInactiveResources || AllApply(conditions, options)) definitions.Add(definition);
            }
        }

        private static bool AllApply(FontFaceCondition[] conditions, HtmlResourcePipelineOptions options) {
            foreach (FontFaceCondition condition in conditions) {
                if (condition.IsMedia ? !IsApplicableMedia(condition.Text, options)
                    : !HtmlComputedStyleEngine.IsApplicableSupports(condition.Text)) return false;
            }
            return true;
        }

        // Mirrors the environment-dependent traversal: a face applies when every enclosing media
        // and supports rule applies, in the same document order.
        private static void CollectFontFaces(
            ICssRule rule,
            List<(HtmlCssFontFaceDefinition, FontFaceCondition[])> faces,
            List<FontFaceCondition> conditions,
            IReadOnlyDictionary<ICssFontFaceRule, IReadOnlyDictionary<string, string>> rawDescriptors) {
            int added = 0;
            if (rule is ICssMediaRule mediaRule) {
                conditions.Add(new FontFaceCondition(true, mediaRule.ConditionText));
                added++;
            }
            if (rule is ICssSupportsRule supportsRule) {
                conditions.Add(new FontFaceCondition(false, supportsRule.ConditionText));
                added++;
            }
            try {
                if (rule is ICssFontFaceRule fontFace) {
                    var definition = new HtmlCssFontFaceDefinition(
                        fontFace.Family,
                        fontFace.Source,
                        fontFace.Weight,
                        fontFace.Stretch,
                        fontFace.Style,
                        fontFace.Range);
                    faces.Add((rawDescriptors.TryGetValue(fontFace, out IReadOnlyDictionary<string, string>? raw)
                        ? definition.WithRawDescriptors(raw)
                        : definition, conditions.ToArray()));
                    return;
                }

                if (rule is ICssGroupingRule groupingRule) {
                    foreach (ICssRule child in groupingRule.Rules) {
                        CollectFontFaces(child, faces, conditions, rawDescriptors);
                    }
                }
            } finally {
                conditions.RemoveRange(conditions.Count - added, added);
            }
        }
    }

    private readonly struct FontFaceCondition {
        internal FontFaceCondition(bool isMedia, string text) {
            IsMedia = isMedia;
            Text = text;
        }

        internal bool IsMedia { get; }
        internal string Text { get; }
    }

    private static readonly FontFaceSheetCache FontFaceSheets = new();

    /// <summary>Thread-safe, least-recently-used parsed sheets, bounded by the text they retain.</summary>
    private sealed class FontFaceSheetCache {
        // Small sheets parse quickly; only sheets that carry font or image payloads are worth keeping.
        private const int MinimumLength = 64 * 1024;
        private const long MaximumRetainedCharacters = 32L * 1024 * 1024;
        private readonly Dictionary<string, LinkedListNode<(string Css, FontFaceSheet Sheet)>> _entries = new(StringComparer.Ordinal);
        private readonly LinkedList<(string Css, FontFaceSheet Sheet)> _order = new();
        private long _retainedCharacters;

        internal FontFaceSheet GetOrParse(string css, HtmlConversionLimits? limits) {
            if (css.Length < MinimumLength || css.Length > MaximumRetainedCharacters) return FontFaceSheet.Parse(css, limits);
            lock (_entries) {
                if (_entries.TryGetValue(css, out LinkedListNode<(string Css, FontFaceSheet Sheet)>? node)) {
                    _order.Remove(node);
                    _order.AddFirst(node);
                    return node.Value.Sheet;
                }
            }
            // A scan that exceeds this operation's limits throws here and remembers nothing.
            FontFaceSheet sheet = FontFaceSheet.Parse(css, limits);
            lock (_entries) {
                if (_entries.ContainsKey(css)) return sheet;
                while (_retainedCharacters + css.Length > MaximumRetainedCharacters && _order.Last != null) {
                    LinkedListNode<(string Css, FontFaceSheet Sheet)> oldest = _order.Last;
                    _order.RemoveLast();
                    _entries.Remove(oldest.Value.Css);
                    _retainedCharacters -= oldest.Value.Css.Length;
                }
                _entries.Add(css, _order.AddFirst((css, sheet)));
                _retainedCharacters += css.Length;
            }
            return sheet;
        }
    }

    internal static IReadOnlyList<string> ExtractFontFaceUrls(string source) {
        return ExtractCssUrls(source);
    }

    internal static IReadOnlyList<string> ExtractCssUrls(string source) {
        var urls = new List<string>();
        if (string.IsNullOrWhiteSpace(source)) {
            return urls.AsReadOnly();
        }

        foreach (Match match in CssUrlExpression.Matches(source)) {
            if (!IsValidCssUrlMatch(source, match) || !IsCssFunctionNameAt(source, match.Index, "url") || IsInsideCssString(source, match.Index)) {
                continue;
            }

            string value = DecodeCssEscapes(match.Groups["url"].Value.Trim().Trim('\'', '"'));
            if (!string.IsNullOrWhiteSpace(value) && !IsFragmentOnlyReference(value)) {
                urls.Add(value);
            }
        }

        return urls.AsReadOnly();
    }

    internal static string RebaseExternalStylesheetUrls(string css, Uri baseUri, HtmlUrlPolicy policy, HtmlResourceSession? resources = null) {
        if (string.IsNullOrWhiteSpace(css)) {
            return css ?? string.Empty;
        }

        var replacements = new List<(int Start, int Length, string Value)>();
        HtmlUrlPolicy resourcePolicy = HtmlResourceUrlPolicy.Create(policy);
        foreach (Match match in CssUrlExpression.Matches(css)) {
            if (!IsValidCssUrlMatch(css, match) || !IsCssFunctionNameAt(css, match.Index, "url") || IsInsideCssString(css, match.Index)) {
                continue;
            }

            string source = DecodeCssEscapes(match.Groups["url"].Value.Trim().Trim('\'', '"'));
            if (string.IsNullOrWhiteSpace(source) || IsFragmentOnlyReference(source)) {
                continue;
            }

            // Imported text may already contain a policy-transformed absolute URL. Its original plan
            // remains authoritative; re-signing it while rebasing an outer sheet changes its identity.
            string resolved = resources != null && resources.TryGetStylesheetImagePlan(source, baseUri, out HtmlResourceReference planned)
                ? HtmlResourceSession.GetMaterializedImageSource(planned)
                : HtmlUrlPolicyEvaluator.ResolveUrl(source, baseUri, resourcePolicy);
            string replacement = resolved.Length == 0
                ? "url(\"\")"
                : "url(\"" + EscapeCssString(resolved) + "\")";
            replacements.Add((match.Index, match.Length, replacement));
        }

        if (replacements.Count == 0) {
            return css;
        }

        var builder = new StringBuilder(css);
        for (int index = replacements.Count - 1; index >= 0; index--) {
            (int start, int length, string value) = replacements[index];
            builder.Remove(start, length);
            builder.Insert(start, value);
        }

        return builder.ToString();
    }

    // An at-keyword spells font-face either literally or through an escape in its name.
    private static bool MayContainFontFaceRule(string css) {
        const string RuleName = "font-face";
        for (int at = css.IndexOf('@'); at >= 0; at = css.IndexOf('@', at + 1)) {
            int name = at + 1;
            if (string.Compare(css, name, RuleName, 0, RuleName.Length, StringComparison.OrdinalIgnoreCase) == 0) {
                return true;
            }

            while (name < css.Length && (char.IsLetterOrDigit(css[name]) || css[name] == '-' || css[name] == '_' || css[name] >= 0x80)) {
                name++;
            }

            if (name < css.Length && css[name] == '\\') {
                return true;
            }
        }

        return false;
    }

    private static IReadOnlyList<IReadOnlyDictionary<string, string>> ExtractRawFontFaceDescriptors(string css, IReadOnlyDictionary<int, int> closures) {
        var rawRules = new List<IReadOnlyDictionary<string, string>>();
        foreach (KeyValuePair<int, int> closure in closures.OrderBy(item => item.Key)) {
            if (!HtmlCssRuleBlockScanner.IsFontFaceBlock(css, closure.Key)) continue;
            string body = css.Substring(closure.Key + 1, closure.Value - closure.Key - 1);
            rawRules.Add(HtmlRenderCssValues.ParseInlineStyleDeclarations(body));
        }
        return rawRules.AsReadOnly();
    }

    private static void MapRawFontFaceDescriptors(
        ICssRule rule,
        IReadOnlyList<IReadOnlyDictionary<string, string>> rawRules,
        IDictionary<ICssFontFaceRule, IReadOnlyDictionary<string, string>> mapped,
        ref int rawRuleIndex) {
        if (rule is ICssFontFaceRule fontFace) {
            while (rawRuleIndex < rawRules.Count) {
                IReadOnlyDictionary<string, string> raw = rawRules[rawRuleIndex++];
                if (!raw.TryGetValue("font-family", out string? rawFamily)
                    || !string.Equals(CleanFontFamily(rawFamily), CleanFontFamily(fontFace.Family), StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }
                mapped[fontFace] = raw;
                break;
            }
            return;
        }
        if (rule is ICssGroupingRule groupingRule) {
            foreach (ICssRule child in groupingRule.Rules) {
                MapRawFontFaceDescriptors(child, rawRules, mapped, ref rawRuleIndex);
            }
        }
    }

    private static string CleanFontFamily(string value) {
        string family = (value ?? string.Empty).Trim();
        while (family.Length >= 2
               && ((family[0] == '"' && family[family.Length - 1] == '"')
                   || (family[0] == '\'' && family[family.Length - 1] == '\''))) {
            family = family.Substring(1, family.Length - 2).Trim();
        }
        return family;
    }

    private static string EscapeCssString(string value) {
        var builder = new StringBuilder(value.Length);
        foreach (char character in value) {
            if (character == '\\' || character == '"' || character == '\r' || character == '\n' || character == '\f') {
                builder.Append('\\').Append(character);
            } else {
                builder.Append(character);
            }
        }

        return builder.ToString();
    }
}
