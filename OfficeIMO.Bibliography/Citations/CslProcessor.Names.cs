using System.Xml.Linq;

namespace OfficeIMO.Bibliography;

public sealed partial class CslProcessor {
    private int NecessaryNameCount(CslRecord[] unresolved, CslRecord[] original, IReadOnlyDictionary<string, int> baseline, CslEvaluator evaluator, XElement layout, IReadOnlyList<CslDisambiguationForm> forms) {
        var unresolvedSet = new HashSet<CslRecord>(unresolved);
        CslRecord[] outside = original.Where(record => !unresolvedSet.Contains(record)).ToArray();
        if (outside.Length == 0) return 0;
        int maximum = unresolved.Max(record => record.MinimumNames);
        Dictionary<string, int> saved = original.ToDictionary(record => record.Key, record => record.MinimumNames, StringComparer.Ordinal);
        try {
            for (int count = 0; count <= maximum; count++) {
                foreach (CslRecord record in original) record.MinimumNames = Math.Max(baseline[record.Key], count);
                var rendered = original.ToDictionary(record => record.Key, record => DisambiguationTexts(record, forms, evaluator, layout).ToArray(), StringComparer.Ordinal);
                var outsideValues = new HashSet<string>(outside.SelectMany(record => rendered[record.Key]), StringComparer.Ordinal);
                if (unresolved.All(record => rendered[record.Key].All(value => !outsideValues.Contains(value)))) return count;
            }
            return maximum;
        } finally { foreach (CslRecord record in original) record.MinimumNames = saved[record.Key]; }
    }

    private void ExpandNames(CslRecord[] records, IReadOnlyList<CslRecord[]> collisions, CslEvaluator evaluator, XElement layout, XElement citation, IReadOnlyList<CslDisambiguationForm> forms) {
        string rule = (string?)citation.Attribute("givenname-disambiguation-rule") ?? "by-cite";
        bool byCite = rule == "by-cite";
        bool primary = rule == "primary-name" || rule == "primary-name-with-initials";
        bool initialsOnly = rule == "all-names-with-initials" || rule == "primary-name-with-initials";
        IEnumerable<CslRecord[]> sets = byCite ? collisions : new[] { records };
        foreach (CslRecord[] set in sets) {
            var occurrences = new List<CslNameOccurrence>();
            foreach (CslRecord record in set) foreach (CslDisambiguationForm form in forms) {
                CslContext context = CreateDisambiguationContext(record, form);
                evaluator.Evaluate(layout, context);
                occurrences.AddRange(context.ObservedNames.Where(name => !primary || name.Primary));
            }
            // Examine name positions in display order. A by-cite rule stops as
            // soon as the current references become distinguishable.
            foreach (IGrouping<string, CslNameOccurrence> group in occurrences.GroupBy(name => name.Rendered, StringComparer.Ordinal)) {
                evaluator.ThrowIfCancellationRequested();
                CslNameOccurrence[] names = group.Where(name => !initialsOnly || evaluator.CanExpandWithInitials(name)).ToArray();
                Dictionary<CslNameOccurrence, string> identities = names.ToDictionary(name => name, evaluator.NameIdentity);
                if (identities.Values.Distinct(StringComparer.Ordinal).Count() < 2) continue;
                var saved = names.Select(name => (name, level: name.Record.NameExpansions.TryGetValue(name.Key, out int level) ? level : 0)).ToArray();
                bool distinct = false;
                for (int level = 1; level <= (initialsOnly ? 1 : 2); level++) {
                    var expanded = names.Select(name => new { Name = name, Text = evaluator.PreviewExpandedName(name, level) }).ToArray();
                    distinct = expanded.GroupBy(value => value.Text, StringComparer.Ordinal).All(values => values.Select(value => identities[value.Name]).Distinct(StringComparer.Ordinal).Count() == 1);
                    if (distinct) break;
                }
                if (!distinct) foreach (var entry in saved) {
                    if (entry.level == 0) entry.name.Record.NameExpansions.Remove(entry.name.Key);
                    else entry.name.Record.NameExpansions[entry.name.Key] = entry.level;
                }
                if (byCite && FindAmbiguities(set, forms, evaluator, layout).Count == 0) break;
            }
        }
    }
}
