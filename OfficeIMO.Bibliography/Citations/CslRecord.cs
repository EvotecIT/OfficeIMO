using System.Text.Json;

namespace OfficeIMO.Bibliography;

internal sealed class CslRecord {
    internal CslRecord(JsonElement data, int number) {
        Data = data.Clone(); Number = number;
        Key = Scalar("id"); Type = Scalar("type");
    }
    internal CslRecord(CslRecord source) { Data = source.Data; Number = source.Number; Key = source.Key; Type = source.Type; }
    internal JsonElement Data { get; }
    internal string Key { get; }
    internal string Type { get; }
    internal int Number { get; set; }
    internal int FirstNote { get; set; }
    internal string YearSuffix { get; set; } = string.Empty;
    internal ISet<string> ActiveConditions { get; } = new HashSet<string>(StringComparer.Ordinal);
    internal int MinimumNames { get; set; }
    internal IDictionary<string, int> NameExpansions { get; } = new Dictionary<string, int>(StringComparer.Ordinal);
    internal JsonElement Value(string variable) {
        // The CSL-JSON input schema also accepts these short-field spellings.
        // A supplied canonical property, including an empty value, retains precedence.
        if (Data.TryGetProperty(variable, out JsonElement value)) return value;
        string? alias = variable == "title-short" ? "shortTitle" : variable == "container-title-short" ? "journalAbbreviation" : null;
        return alias != null && Data.TryGetProperty(alias, out value) && value.ValueKind == JsonValueKind.String ? value : default;
    }
    internal string Scalar(string variable) {
        JsonElement value = Value(variable);
        return value.ValueKind == JsonValueKind.String ? value.GetString() ?? string.Empty : value.ValueKind == JsonValueKind.Number ?
            variable == "id" && value.TryGetDecimal(out decimal number) ? number.ToString("G29", CultureInfo.InvariantCulture) : value.GetRawText() : string.Empty;
    }
    internal JsonElement[] Names(string variable, CancellationToken token = default) {
        JsonElement value = Value(variable);
        if (variable == "editor-translator" && value.ValueKind == JsonValueKind.Undefined) {
            JsonElement[] editors = Names("editor"), translators = Names("translator");
            return SameNames(editors, translators, token) ? editors : Array.Empty<JsonElement>();
        }
        return value.ValueKind == JsonValueKind.Array ? value.EnumerateArray().Where(name => name.ValueKind == JsonValueKind.Object).ToArray() : Array.Empty<JsonElement>();
    }
    internal static bool SameNames(JsonElement[] left, JsonElement[] right, CancellationToken token) => left.Length > 0 && left.Length == right.Length &&
        left.Select(name => CslNameIdentity.Read(name, true, token)).SequenceEqual(right.Select(name => CslNameIdentity.Read(name, true, token)), StringComparer.Ordinal);
    internal bool Has(string variable) {
        JsonElement value = Value(variable);
        return value.ValueKind != JsonValueKind.Undefined && value.ValueKind != JsonValueKind.Null &&
            (value.ValueKind != JsonValueKind.Array || value.GetArrayLength() > 0) &&
            (value.ValueKind != JsonValueKind.String || !string.IsNullOrEmpty(value.GetString()));
    }
}

internal sealed class CslContext {
    internal CslContext(CslRecord record, XElementScope scope) { Record = record; Scope = scope; }
    internal CslRecord Record { get; }
    internal XElementScope Scope { get; }
    internal CslCitationItem? Cite { get; set; }
    internal bool Sorting { get; set; }
    internal string Position { get; set; } = "first";
    internal bool NearNote { get; set; }
    internal int? NoteIndex { get; set; }
    internal string MacroPath { get; set; } = string.Empty;
    // Cite metadata stays fixed during a context's evaluation. Reuse its locator
    // classification across repeated conditional branches instead of rescanning it.
    internal string? LocatorDisambiguationKey { get; set; }
    internal IList<string>? ObservedConditions { get; set; }
    internal int? SortNamesMinimum { get; set; }
    internal int? SortNamesFirst { get; set; }
    internal bool? SortNamesLast { get; set; }
    internal string InheritedDelimiter { get; set; } = string.Empty;
    internal bool SuppressFirstNames { get; set; }
    internal bool SuppressYear { get; set; }
    internal bool YearRendered { get; set; }
    internal bool AutomaticYearSuffix { get; set; } = true;
    internal bool FirstNamesHandled { get; set; }
    internal string FirstNamesText { get; set; } = string.Empty;
    internal CslText? NarrativeNames { get; set; }
    internal int NamesDepth { get; set; }
    internal string[] FirstNames { get; set; } = Array.Empty<string>();
    internal string[] PreviousNames { get; set; } = Array.Empty<string>();
    internal string? NamesReplacement { get; set; }
    internal string NamesReplacementRule { get; set; } = "complete-all";
    internal bool CompleteNamesMatch { get; set; }
    // An intentionally blank author replacement still selects its fallback.
    internal int EmptyNamesReplacementCount { get; set; }
    internal IList<CslNameOccurrence> ObservedNames { get; } = new List<CslNameOccurrence>();
    internal ISet<string> Suppressed { get; } = new HashSet<string>(StringComparer.Ordinal);
    internal ISet<string> RenderedVariables { get; } = new HashSet<string>(StringComparer.Ordinal);
    private CslSubstitutionScope? _substitution;
    internal void RegisterRenderedVariable(string variable) {
        RenderedVariables.Add(variable);
        if (_substitution != null && !_substitution.PreviouslyRendered.Contains(variable)) Suppressed.Add(variable);
    }
    internal CslSubstitutionScope BeginSubstitution() {
        var scope = new CslSubstitutionScope(_substitution, this);
        _substitution = scope;
        return scope;
    }
    internal void EndSubstitution(CslSubstitutionScope scope, bool successful) {
        _substitution = scope.Parent;
        if (successful) return;
        RenderedVariables.Clear();
        RenderedVariables.UnionWith(scope.PreviouslyRendered);
        Suppressed.Clear();
        Suppressed.UnionWith(scope.PreviouslySuppressed);
        FirstNamesHandled = scope.FirstNamesHandled;
        FirstNames = scope.FirstNames;
        FirstNamesText = scope.FirstNamesText;
        CompleteNamesMatch = scope.CompleteNamesMatch;
        EmptyNamesReplacementCount = scope.EmptyNamesReplacementCount;
        while (ObservedNames.Count > scope.ObservedNamesCount) ObservedNames.RemoveAt(ObservedNames.Count - 1);
        NarrativeNames = scope.NarrativeNames;
        YearRendered = scope.YearRendered;
    }
    internal string Variable(string name, bool includeSuppressed = false) {
        if (!includeSuppressed && Suppressed.Contains(name)) return string.Empty;
        switch (name) {
            case "citation-number": return Record.Number.ToString(CultureInfo.InvariantCulture);
            case "year-suffix": return Record.YearSuffix;
            case "first-reference-note-number": return Record.FirstNote == 0 || Scope == XElementScope.Citation &&
                (Position == "first" && !Sorting || NoteIndex.HasValue && (NoteIndex.Value == 0 || NoteIndex.Value <= Record.FirstNote))
                ? string.Empty : Record.FirstNote.ToString(CultureInfo.InvariantCulture);
            case "locator": return Cite?.Locator ?? string.Empty;
            case "citation-label": return Record.Scalar(name);
            case "page-first":
                string page = Record.Scalar("page-first");
                if (page.Length > 0) return page;
                return Record.Scalar("page").Split(new[] { '-', '–', ',', '&' })[0].Trim();
            default: return Record.Scalar(name);
        }
    }
}

internal enum XElementScope { Citation, Bibliography }

internal sealed class CslNameOccurrence {
    internal CslNameOccurrence(CslRecord record, string variable, int index, JsonElement data, System.Xml.Linq.XElement name, string form, string rendered, bool primary) {
        Record = record; Variable = variable; Index = index; Data = data; Element = name; Form = form; Rendered = rendered; Primary = primary;
    }
    internal CslRecord Record { get; }
    internal string Variable { get; }
    internal int Index { get; }
    internal string Key => Variable + "\u001f" + Index.ToString(CultureInfo.InvariantCulture);
    internal JsonElement Data { get; }
    internal System.Xml.Linq.XElement Element { get; }
    internal string Form { get; }
    internal string Rendered { get; }
    internal bool Primary { get; }
}
