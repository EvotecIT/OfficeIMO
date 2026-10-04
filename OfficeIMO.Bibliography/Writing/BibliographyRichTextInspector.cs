namespace OfficeIMO.Bibliography;

/// <summary>Separates source-preserving text storage from destination formatting semantics.</summary>
internal static class BibliographyRichTextInspector {
    internal static void Inspect(BibliographyItem item, BibliographyFormat source, BibliographyFormat target, BibliographyConversionReport report, CancellationToken token) {
        if (source == target || source != BibliographyFormat.CslJson && source != BibliographyFormat.BibTex && source != BibliographyFormat.BibLatex) return;
        string?[] values = { item.Title, item.ContainerTitle, item.CollectionTitle, item.Abstract };
        string[] fields = { "title", "container-title", "collection-title", "abstract" };
        for (int index = 0; index < values.Length; index++) {
            token.ThrowIfCancellationRequested();
            if (values[index] != null) InspectValue(values[index]!, fields[index]);
        }
        foreach (string note in item.Notes) { token.ThrowIfCancellationRequested(); InspectValue(note, "notes"); }

        void InspectValue(string value, string field) {
            bool semantic = source == BibliographyFormat.CslJson ? HasCslMarkup(value, token) : HasBibMarkup(value, token);
            if (!semantic || source != BibliographyFormat.CslJson && (target == BibliographyFormat.BibTex || target == BibliographyFormat.BibLatex)) return;
            report.Add("BIBCONV252", BibliographyDiagnosticSeverity.Warning,
                $"Rich text or case-protection syntax in '{field}' is retained literally in {target}; its source formatting semantics cannot be reopened exactly.",
                BibliographyConversionAction.Approximated, item, field);
        }
    }

    private static bool HasCslMarkup(string value, CancellationToken token) {
        for (int index = 0; index < value.Length; index++) {
            if ((index & 4095) == 0) token.ThrowIfCancellationRequested();
            if (value[index] == '<') {
                int start = index + 1;
                if (start < value.Length && value[start] == '/') start++;
                int end = start;
                while (end < value.Length && end - start < 8 && char.IsLetter(value[end])) end++;
                if (end >= value.Length || value[end] != '>' && !char.IsWhiteSpace(value[end])) continue;
                string name = value.Substring(start, end - start);
                if (name == "i" || name == "b" || name == "sup" || name == "sub" || name == "sc" || name == "span") return true;
            } else if (value[index] == '&') {
                int end = index + 1;
                while (end < value.Length && end - index <= 32 && value[end] != ';' && !char.IsWhiteSpace(value[end])) end++;
                if (end < value.Length && value[end] == ';') {
                    string entity = value.Substring(index, end - index + 1);
                    if (!string.Equals(System.Net.WebUtility.HtmlDecode(entity), entity, StringComparison.Ordinal)) return true;
                }
            }
        }
        return false;
    }

    private static bool HasBibMarkup(string value, CancellationToken token) {
        for (int index = 0; index < value.Length; index++) {
            if ((index & 4095) == 0) token.ThrowIfCancellationRequested();
            if (value[index] == '{' || value[index] == '}' || value[index] == '\\') return true;
        }
        return false;
    }
}
