using System.Text.Json;
using System.Xml.Linq;

namespace OfficeIMO.Bibliography;

public sealed partial class CslProcessor {
    private void Disambiguate(CslRecord[] records, IReadOnlyList<CslCitation> clusters, CslEvaluator evaluator) {
        XElement citation = _style.Root.Element(CslStyle.Namespace + "citation")!;
        XElement layout = CslElementIdentity.Copy(citation.Element(CslStyle.Namespace + "layout")!);
        layout.Attribute("prefix")?.Remove(); layout.Attribute("suffix")?.Remove();
        // Full notes can identify a work while its shortened note remains
        // ambiguous. Compare each reusable form before adding global suffixes.
        var forms = new List<CslDisambiguationForm> { new CslDisambiguationForm("first") };
        if (_style.HasSubsequentForm) {
            forms.Add(new CslDisambiguationForm("subsequent"));
            if (_style.IsNoteStyle && _style.HasNearNoteCondition)
                forms.Add(new CslDisambiguationForm("subsequent", true));
        }
        if (_style.HasLocatorConditions) {
            CslDisambiguationForm[] positions = forms.ToArray();
            var locators = new HashSet<(string Type, bool Numeric)>();
            foreach (CslCitation cluster in clusters) foreach (CslCitationItem item in cluster.Items) {
                evaluator.PerformOperation();
                if (string.IsNullOrEmpty(item.Locator)) continue;
                bool numeric = CslNumberSyntax.IsNumeric(item.Locator!, evaluator.CancellationToken);
                if (!locators.Add((item.LocatorType, numeric))) continue;
                foreach (CslDisambiguationForm position in positions)
                    forms.Add(new CslDisambiguationForm(position.Position, position.NearNote, item.LocatorType, numeric));
            }
        }
        List<CslRecord[]> Collisions() => ConnectedAmbiguities(records,
            FindAmbiguities(records, forms, evaluator, layout).Select(ambiguity => ambiguity.Records), evaluator);
        List<CslRecord[]> collisions = Collisions();
        bool addGiven = (string?)citation.Attribute("disambiguate-add-givenname") == "true";
        if (addGiven) { ExpandNames(records, collisions, evaluator, layout, citation, forms); collisions = Collisions(); }
        if ((string?)citation.Attribute("disambiguate-add-names") == "true") {
            Dictionary<string, int> originalCounts = records.ToDictionary(record => record.Key, record => record.MinimumNames, StringComparer.Ordinal);
            var originalGroups = collisions.SelectMany(group => group.Select(record => (Record: record, Group: group)))
                .ToDictionary(entry => entry.Record, entry => entry.Group);
            int maximumNames = records.Length == 0 ? 0 : records.Max(record => record.Data.EnumerateObject()
                .Where(property => property.Value.ValueKind == JsonValueKind.Array).Select(property => property.Value.GetArrayLength()).DefaultIfEmpty().Max());
            for (int minimum = 2; minimum <= maximumNames && collisions.Count > 0; minimum++) {
                foreach (CslRecord record in collisions.SelectMany(group => group)) record.MinimumNames = Math.Max(record.MinimumNames, minimum);
                if (addGiven) ExpandNames(records, collisions, evaluator, layout, citation, forms);
                collisions = Collisions();
            }
            foreach (CslRecord[] group in collisions) {
                int needed = originalGroups.TryGetValue(group[0], out CslRecord[]? original)
                    ? NecessaryNameCount(group, original, originalCounts, evaluator, layout, forms) : 0;
                foreach (CslRecord record in group) record.MinimumNames = Math.Max(originalCounts[record.Key], needed);
            }
        }
        if (_style.HasConditionalDisambiguation) SelectConditionalDetail(records, forms, evaluator, layout);
        if ((string?)citation.Attribute("disambiguate-add-year-suffix") == "true") {
            bool assigned = AssignYearSuffixes(records, forms, evaluator, layout);
            // An explicit suffix guarded by disambiguate is empty in earlier
            // trials. Reconsider detail once suffix assignment can distinguish it.
            if (assigned && _style.HasConditionalDisambiguation) SelectConditionalDetail(records, forms, evaluator, layout, reassignSuffixes: true);
        }
    }

    /// <summary>Assigns suffixes from unsuffixed output after the current conditional choices, preserving bibliography order.</summary>
    private bool AssignYearSuffixes(CslRecord[] records, IReadOnlyList<CslDisambiguationForm> forms, CslEvaluator evaluator, XElement layout) {
        foreach (CslRecord record in records) record.YearSuffix = string.Empty;
        List<CslRecord[]> groups = ConnectedAmbiguities(records, FindAmbiguities(records, forms, evaluator, layout).Select(ambiguity => ambiguity.Records), evaluator);
        foreach (CslRecord[] group in groups) for (int index = 0; index < group.Length; index++) {
            evaluator.PerformOperation();
            group[index].YearSuffix = AlphabeticSuffix(index);
        }
        return groups.Count > 0;
    }

    private List<CslAmbiguity> FindAmbiguities(CslRecord[] records, IReadOnlyList<CslDisambiguationForm> forms, CslEvaluator evaluator, XElement layout) =>
        Ambiguities(DisambiguationValues(records, forms, evaluator, layout, false));

    private static List<CslAmbiguity> Ambiguities(IEnumerable<CslDisambiguationValue> values) => values
        .GroupBy(value => value.Text, StringComparer.Ordinal)
        // Equal outputs of the same work in multiple forms are not ambiguous.
        .Select(group => new CslAmbiguity(group.ToArray())).Where(ambiguity => ambiguity.Records.Length > 1).ToList();

    private IEnumerable<string> DisambiguationTexts(CslRecord record, IReadOnlyList<CslDisambiguationForm> forms, CslEvaluator evaluator, XElement layout) =>
        forms.Select(form => evaluator.Evaluate(layout, CreateDisambiguationContext(record, form)).Plain);

    private CslContext CreateDisambiguationContext(CslRecord record, CslDisambiguationForm form) {
        CslContext context = CreateContext(record, XElementScope.Citation);
        context.Position = form.Position;
        context.NearNote = form.NearNote;
        // Use the same locator value for every work in this form. Different page
        // numbers must not conceal an otherwise ambiguous source identity.
        if (form.LocatorType != null) context.Cite = new CslCitationItem(record.Key) {
            LocatorType = form.LocatorType, Locator = form.NumericLocator ? "1" : "source"
        };
        return context;
    }

    private static List<CslRecord[]> ConnectedAmbiguities(CslRecord[] records, IEnumerable<CslRecord[]> collisions, CslEvaluator evaluator) {
        var parents = new Dictionary<CslRecord, CslRecord>();
        CslRecord Root(CslRecord record) {
            if (!parents.ContainsKey(record)) parents.Add(record, record);
            CslRecord root = record;
            while (parents[root] != root) { evaluator.PerformOperation(); root = parents[root]; }
            while (parents[record] != record) { CslRecord next = parents[record]; parents[record] = root; record = next; }
            return root;
        }
        foreach (CslRecord[] group in collisions) foreach (CslRecord record in group) {
            evaluator.PerformOperation();
            parents[Root(record)] = Root(group[0]);
        }
        return records.Where(parents.ContainsKey).GroupBy(Root).Select(group => group.ToArray()).ToList();
    }

    private sealed class CslDisambiguationForm {
        internal CslDisambiguationForm(string position, bool nearNote = false, string? locatorType = null, bool numericLocator = true) {
            Position = position; NearNote = nearNote; LocatorType = locatorType; NumericLocator = numericLocator;
        }
        internal string Position { get; }
        internal bool NearNote { get; }
        internal string? LocatorType { get; }
        internal bool NumericLocator { get; }
    }

    private sealed class CslDisambiguationValue {
        internal CslDisambiguationValue(CslRecord record, CslDisambiguationForm form, string text, IList<string>? conditions) { Record = record; Form = form; Text = text; Conditions = conditions; }
        internal CslRecord Record { get; }
        internal CslDisambiguationForm Form { get; }
        internal string Text { get; }
        internal IList<string>? Conditions { get; }
    }

    private sealed class CslAmbiguity {
        internal CslAmbiguity(CslDisambiguationValue[] values) { Values = values; Records = values.Select(value => value.Record).Distinct().ToArray(); }
        internal CslDisambiguationValue[] Values { get; }
        internal CslRecord[] Records { get; }
    }
}
