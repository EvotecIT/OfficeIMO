using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionPlanner {
    /// <summary>Derives reviewable redaction rectangles from literal text, bounded regex, logical element kinds, and AcroForm field names.</summary>
    public static PdfRedactionPlan Search(byte[] pdf, PdfRedactionSearchOptions search, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf)); Guard.NotNull(search, nameof(search));
        if (search.RegexTimeout <= TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(search), "Regex timeout must be positive.");
        Regex[] expressions = search.RegularExpressions.Select(pattern => new Regex(pattern, search.RegexOptions, search.RegexTimeout)).ToArray();
        if (search.LiteralText.Count == 0 && expressions.Length == 0 && search.FormFieldNames.Count == 0 && search.LogicalElementKinds.Count == 0) throw new ArgumentException("At least one redaction search criterion is required.", nameof(search));

        PdfLogicalDocument logical = PdfLogicalDocument.From(PdfReadDocument.Open(pdf, readOptions), layoutOptions);
        StringComparison comparison = search.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var areas = new List<PdfRedactionArea>(); var keys = new HashSet<string>(StringComparer.Ordinal);
        foreach (PdfLogicalTextBlock block in logical.TextBlocks) {
            string? criterion = MatchText(block, search, expressions, comparison); if (criterion is null) continue;
            double fontSize = GetEffectiveFontSize(block); double x = Math.Min(block.XStart, block.XEnd); double width = Math.Max(1D, Math.Abs(block.XEnd - block.XStart));
            AddArea(areas, keys, new PdfRedactionArea(block.PageNumber, x, block.BaselineY - fontSize, width, fontSize * 1.5D, criterion));
        }
        var requestedFields = new HashSet<string>(search.FormFieldNames, StringComparer.Ordinal);
        foreach (PdfLogicalFormWidget widget in logical.FormWidgets) if (widget.FieldName is not null && requestedFields.Contains(widget.FieldName)) AddArea(areas, keys, new PdfRedactionArea(widget.PageNumber, widget.X1, widget.Y1, widget.Width, widget.Height, "field:" + widget.FieldName));
        if (areas.Count == 0) return new PdfRedactionPlan(PdfInspector.Preflight(pdf, readOptions), Array.Empty<PdfRedactionArea>(), Array.Empty<PdfRedactionMatch>(), new[] { new PdfDiagnosticFinding(PdfDiagnosticSeverity.Info, "RedactionSearchNoMatches", "No logical content matched the requested redaction search criteria.") }, DescribeCriteria(search));
        PdfRedactionPlan planned = Plan(pdf, areas, layoutOptions, readOptions);
        return new PdfRedactionPlan(planned.Preflight, planned.Areas, planned.Matches, planned.Findings, DescribeCriteria(search));
    }

    private static string? MatchText(PdfLogicalTextBlock block, PdfRedactionSearchOptions search, Regex[] expressions, StringComparison comparison) {
        for (int i = 0; i < search.LiteralText.Count; i++) if (ContainsText(block.Text, search.LiteralText[i], comparison)) return "literal:" + search.LiteralText[i];
        for (int i = 0; i < expressions.Length; i++) if (expressions[i].IsMatch(block.Text)) return "regex:" + search.RegularExpressions[i];
        return search.LogicalElementKinds.Contains(block.Kind) ? "logical-kind:" + block.Kind.ToString() : null;
    }

    private static void AddArea(List<PdfRedactionArea> areas, HashSet<string> keys, PdfRedactionArea area) { string key = area.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" + area.X.ToString("R", System.Globalization.CultureInfo.InvariantCulture) + ":" + area.Y.ToString("R", System.Globalization.CultureInfo.InvariantCulture) + ":" + area.Width.ToString("R", System.Globalization.CultureInfo.InvariantCulture) + ":" + area.Height.ToString("R", System.Globalization.CultureInfo.InvariantCulture); if (keys.Add(key)) areas.Add(area); }
    private static string[] DescribeCriteria(PdfRedactionSearchOptions search) => search.LiteralText.Select(value => "literal:" + value).Concat(search.RegularExpressions.Select(value => "regex:" + value)).Concat(search.FormFieldNames.Select(value => "field:" + value)).Concat(search.LogicalElementKinds.Select(value => "logical-kind:" + value.ToString())).ToArray();
    private static bool ContainsText(string text, string value, StringComparison comparison) { if (value.Length == 0) return true; for (int i = 0; i <= text.Length - value.Length; i++) if (string.Compare(text, i, value, 0, value.Length, comparison) == 0) return true; return false; }
}
