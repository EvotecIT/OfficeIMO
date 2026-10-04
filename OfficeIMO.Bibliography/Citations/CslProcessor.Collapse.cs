using System.Xml.Linq;

namespace OfficeIMO.Bibliography;

public sealed partial class CslProcessor {
    private CslText Collapse(IReadOnlyList<CiteRendering> source, XElement citation, XElement layout, CslEvaluator evaluator) {
        string delimiter = CslEvaluator.Attr(layout, "delimiter") ?? string.Empty;
        string? mode = CslEvaluator.Attr(citation, "collapse");
        if (mode == "citation-number") return CollapseNumbers(source, delimiter);
        if (mode == null && citation.Attribute("cite-group-delimiter") == null) return Join(source.Select(cite => cite.Value), delimiter);

        // Stable grouping preserves the order of both first occurrences and the
        // cites within each author group, after the style's citation sort.
        var groups = new List<List<CiteRendering>>();
        var names = new Dictionary<string, List<CiteRendering>>(StringComparer.Ordinal);
        foreach (CiteRendering cite in source) {
            evaluator.ThrowIfCancellationRequested();
            string key = cite.Context.FirstNamesText;
            // Date-only layouts and empty name expressions share an empty visible name group.
            if (!names.TryGetValue(key, out List<CiteRendering>? group)) {
                group = new List<CiteRendering>(); groups.Add(group);
                names.Add(key, group);
            }
            group.Add(cite);
        }
        var result = CslText.Empty;
        bool precedingCollapsed = false;
        foreach (List<CiteRendering> group in groups) {
            evaluator.ThrowIfCancellationRequested();
            var values = new List<CslText>();
            var separators = new List<string>();
            bool precedingSuffixOnly = false;
            bool groupCollapsed = false;
            for (int index = 0; index < group.Count; index++) {
                evaluator.ThrowIfCancellationRequested();
                CiteRendering cite = group[index];
                CslText value = cite.Value;
                bool collapse = index > 0 && mode != null && !HasAffixes(cite.Item) && !HasAffixes(group[index - 1].Item);
                bool suffixOnly = collapse && (mode == "year-suffix" || mode == "year-suffix-ranged") &&
                    cite.Context.Record.YearSuffix.Length > 0 && group[index - 1].Context.Record.YearSuffix.Length > 0 &&
                    !HasLocator(cite.Item) && !HasLocator(group[index - 1].Item) && SameYear(cite.Context.Record, group[index - 1].Context.Record);
                if (collapse) {
                    groupCollapsed = true;
                    CslContext context = CreateContext(cite.Context.Record, XElementScope.Citation);
                    context.Cite = cite.Item; context.Position = cite.Context.Position; context.NearNote = cite.Context.NearNote;
                    context.NoteIndex = cite.Context.NoteIndex;
                    context.SuppressFirstNames = true; context.SuppressYear = suffixOnly;
                    value = evaluator.Evaluate(ItemLayout(layout), context).Affix(cite.Item.Prefix ?? string.Empty, cite.Item.Suffix ?? string.Empty);
                    if (suffixOnly) value = value.TrimStart(evaluator.CancellationToken);
                }
                values.Add(value);
                separators.Add(suffixOnly ? CslEvaluator.Attr(citation, "year-suffix-delimiter") ?? delimiter :
                    precedingSuffixOnly || mode != null && (HasLocator(cite.Item) || index > 0 && HasLocator(group[index - 1].Item)) ?
                    CslEvaluator.Attr(citation, "after-collapse-delimiter") ?? delimiter :
                    CslEvaluator.Attr(citation, "cite-group-delimiter") ?? ", ");
                precedingSuffixOnly = suffixOnly;
            }
            if (mode == "year-suffix-ranged") CollapseSuffixes(group, values, separators);
            CslText rendered = CslText.Empty;
            for (int index = 0; index < values.Count; index++) rendered = Join(new[] { rendered, values[index] }, separators[index]);
            result = Join(new[] { result, rendered }, precedingCollapsed ? CslEvaluator.Attr(citation, "after-collapse-delimiter") ?? delimiter : delimiter);
            precedingCollapsed = groupCollapsed;
        }
        return result;
    }

    private CslText CollapseNumbers(IReadOnlyList<CiteRendering> source, string delimiter) {
        var values = new List<CslText>();
        for (int index = 0; index < source.Count; index++) {
            int end = index;
            while (end + 1 < source.Count && !HasAffixes(source[end].Item) && !HasAffixes(source[end + 1].Item) &&
                !HasLocator(source[end].Item) && !HasLocator(source[end + 1].Item) &&
                source[end + 1].Context.Record.Number == source[end].Context.Record.Number + 1) end++;
            if (end - index >= 2) {
                values.Add(Join(new[] { source[index].Value, source[end].Value }, "–")); index = end;
            } else values.Add(source[index].Value);
        }
        return Join(values, delimiter);
    }

    private void CollapseSuffixes(IReadOnlyList<CiteRendering> group, IList<CslText> values, IList<string> separators) {
        for (int index = 0; index < group.Count; index++) {
            int end = index;
            while (end + 1 < group.Count && !HasLocator(group[end].Item) && !HasLocator(group[end + 1].Item) &&
                !HasAffixes(group[end].Item) && !HasAffixes(group[end + 1].Item) && SameYear(group[end].Context.Record, group[end + 1].Context.Record) &&
                SuffixNumber(group[end + 1].Context.Record.YearSuffix) == SuffixNumber(group[end].Context.Record.YearSuffix) + 1) end++;
            if (end - index < 2) continue;
            values[index] = Join(new[] { values[index], values[end] }, "–");
            for (int removed = index + 1; removed <= end; removed++) values[removed] = CslText.Empty;
            index = end;
        }
    }

    private static bool SameYear(CslRecord left, CslRecord right) {
        string Year(CslRecord record) {
            var date = record.Value("issued");
            return date.ValueKind == System.Text.Json.JsonValueKind.Object && date.TryGetProperty("date-parts", out var parts) &&
                parts.ValueKind == System.Text.Json.JsonValueKind.Array && parts.GetArrayLength() > 0 && parts[0].ValueKind == System.Text.Json.JsonValueKind.Array && parts[0].GetArrayLength() > 0 ? parts[0][0].ToString() : string.Empty;
        }
        string year = Year(left);
        return year.Length > 0 && year == Year(right);
    }
    private static long SuffixNumber(string value) {
        if (value.Length == 0 || value.Length > 6) return -2;
        long number = 0;
        foreach (char letter in value) { if (letter < 'a' || letter > 'z') return -2; number = number * 26 + letter - 'a' + 1; }
        return number;
    }
    private static bool HasLocator(CslCitationItem item) => !string.IsNullOrEmpty(item.Locator);
    private static bool HasAffixes(CslCitationItem item) => !string.IsNullOrEmpty(item.Prefix) || !string.IsNullOrEmpty(item.Suffix) || item.AuthorOnly || item.SuppressAuthor;
    private sealed class CiteRendering {
        internal CiteRendering(CslCitationItem item, CslContext context, CslText value) { Item = item; Context = context; Value = value; }
        internal CslCitationItem Item { get; }
        internal CslContext Context { get; }
        internal CslText Value { get; }
    }
}
