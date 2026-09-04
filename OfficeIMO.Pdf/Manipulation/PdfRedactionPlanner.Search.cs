using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionPlanner {
    /// <summary>Derives reviewable redaction rectangles from literal text, bounded regex, logical element kinds, and AcroForm field names.</summary>
    public static PdfRedactionPlan Search(byte[] pdf, PdfRedactionSearchOptions search, PdfTextLayoutOptions? layoutOptions = null, PdfLoadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf)); Guard.NotNull(search, nameof(search));
        search.CancellationToken.ThrowIfCancellationRequested();
        if (search.RegexTimeout <= TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(search), "Regex timeout must be positive.");
        if (search.MaximumCandidates <= 0) throw new ArgumentOutOfRangeException(nameof(search), "Maximum candidates must be positive.");
        Regex[] expressions = search.RegularExpressions.Select(pattern => new Regex(pattern, search.RegexOptions, search.RegexTimeout)).ToArray();
        if (search.LiteralText.Count == 0 && expressions.Length == 0 && search.FormFieldNames.Count == 0 && search.LogicalElementKinds.Count == 0) throw new ArgumentException("At least one redaction search criterion is required.", nameof(search));

        PdfReadDocument readDocument = PdfReadDocument.Open(pdf, readOptions, search.CancellationToken);
        PdfDocumentReadResult logical = PdfDocumentReadResult.From(readDocument, layoutOptions);
        search.CancellationToken.ThrowIfCancellationRequested();
        StringComparison comparison = search.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var areas = new List<PdfRedactionArea>(); var keys = new HashSet<string>(StringComparer.Ordinal);
        foreach (PdfLogicalTextBlock block in logical.TextBlocks) {
            search.CancellationToken.ThrowIfCancellationRequested();
            string? criterion = MatchText(block, search, expressions, comparison); if (criterion is null) continue;
            PdfTextSpanBounds bounds = GetTextBlockBounds(block, logical.Pages[block.PageNumber - 1]);
            AddArea(areas, keys, new PdfRedactionArea(block.PageNumber, bounds.Left, bounds.Bottom, bounds.Width, bounds.Height, criterion), search.MaximumCandidates);
        }
        var requestedFields = new HashSet<string>(search.FormFieldNames, StringComparer.Ordinal);
        foreach (PdfLogicalFormWidget widget in logical.FormWidgets) {
            search.CancellationToken.ThrowIfCancellationRequested();
            if (widget.FieldName is not null && requestedFields.Contains(widget.FieldName)) AddArea(areas, keys, new PdfRedactionArea(widget.PageNumber, widget.X1, widget.Y1, widget.Width, widget.Height, "field:" + widget.FieldName), search.MaximumCandidates);
        }
        if (areas.Count == 0) return new PdfRedactionPlan(PdfInspector.Preflight(pdf, readOptions, search.CancellationToken), Array.Empty<PdfRedactionArea>(), Array.Empty<PdfRedactionMatch>(), new[] { new PdfDiagnosticFinding(PdfDiagnosticSeverity.Info, "RedactionSearchNoMatches", "No logical content matched the requested redaction search criteria.") }, DescribeCriteria(search), PdfRedactionPlan.ComputeSourceSha256(pdf), PdfRedactionPlan.CapturePageIdentities(readDocument, Array.Empty<PdfRedactionArea>()));
        PdfRedactionPlan planned = Plan(pdf, areas, layoutOptions, readOptions, search.CancellationToken);
        search.CancellationToken.ThrowIfCancellationRequested();
        return new PdfRedactionPlan(planned.Preflight, planned.Areas, planned.Matches, planned.Findings, DescribeCriteria(search), planned.SourceSha256, planned.PageIdentities, planned.ReviewedTextObjectScopes);
    }

    private static string? MatchText(PdfLogicalTextBlock block, PdfRedactionSearchOptions search, Regex[] expressions, StringComparison comparison) {
        for (int i = 0; i < search.LiteralText.Count; i++) if (ContainsText(block.Text, search.LiteralText[i], comparison)) return "literal:" + search.LiteralText[i];
        for (int i = 0; i < expressions.Length; i++) if (expressions[i].IsMatch(block.Text)) return "regex:" + search.RegularExpressions[i];
        return search.LogicalElementKinds.Contains(block.Kind) ? "logical-kind:" + block.Kind.ToString() : null;
    }

    private static void AddArea(List<PdfRedactionArea> areas, HashSet<string> keys, PdfRedactionArea area, int maximumCandidates) { string key = area.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" + area.X.ToString("R", System.Globalization.CultureInfo.InvariantCulture) + ":" + area.Y.ToString("R", System.Globalization.CultureInfo.InvariantCulture) + ":" + area.Width.ToString("R", System.Globalization.CultureInfo.InvariantCulture) + ":" + area.Height.ToString("R", System.Globalization.CultureInfo.InvariantCulture); if (keys.Add(key)) { if (areas.Count >= maximumCandidates) throw new InvalidOperationException("Redaction search exceeded the configured candidate limit."); areas.Add(area); } }
    private static string[] DescribeCriteria(PdfRedactionSearchOptions search) => search.LiteralText.Select(value => "literal:" + value).Concat(search.RegularExpressions.Select(value => "regex:" + value)).Concat(search.FormFieldNames.Select(value => "field:" + value)).Concat(search.LogicalElementKinds.Select(value => "logical-kind:" + value.ToString())).ToArray();
    private static bool ContainsText(string text, string value, StringComparison comparison) { if (value.Length == 0) return true; for (int i = 0; i <= text.Length - value.Length; i++) if (string.Compare(text, i, value, 0, value.Length, comparison) == 0) return true; return false; }
}
