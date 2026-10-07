using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionPlanner {
    private static PdfRedactionPlan SearchMatchedGlyphs(byte[] pdf, PdfReadDocument readDocument,
        PdfRedactionSearchOptions search, Regex[] expressions, PdfTextLayoutOptions? layoutOptions, PdfLoadOptions? readOptions) {
        var budget = new PdfRedactionSearchWorkBudget("matched-glyph search planning");
        IReadOnlyList<PdfRedactionArea> textAreas = PdfTextEditor.FindRedactionAreas(pdf, search, expressions, readOptions, budget,
            out IReadOnlyList<PdfDiagnosticFinding> selectionFindings);
        var areas = new List<PdfRedactionArea>();
        var keys = new HashSet<string>(StringComparer.Ordinal);
        foreach (PdfRedactionArea area in textAreas) AddArea(areas, keys, area, search.MaximumCandidates, search.ContentScope);
        if (search.FormFieldNames.Count > 0) {
            var fields = new HashSet<string>(search.FormFieldNames, StringComparer.Ordinal);
            PdfDocumentReadResult logical = PdfDocumentReadResult.From(readDocument, layoutOptions, search.CancellationToken);
            foreach (PdfLogicalFormWidget widget in logical.FormWidgets) {
                search.CancellationToken.ThrowIfCancellationRequested();
                if (search.PageNumbers.Count > 0 && !search.PageNumbers.Contains(widget.PageNumber)) continue;
                if (widget.FieldName is not null && fields.Contains(widget.FieldName)) AddArea(areas, keys,
                    new PdfRedactionArea(widget.PageNumber, widget.X1, widget.Y1, widget.Width, widget.Height, "field:" + widget.FieldName),
                    search.MaximumCandidates, search.ContentScope);
            }
        }
        if (areas.Count == 0) {
            PdfDiagnosticFinding[] findings = selectionFindings.Count > 0 ? selectionFindings.ToArray() : new[] {
                new PdfDiagnosticFinding(PdfDiagnosticSeverity.Info, "RedactionSearchNoMatches", "No native text or form widgets matched the requested criteria.")
            };
            return new PdfRedactionPlan(PdfInspector.Preflight(pdf, readOptions, search.CancellationToken), areas.AsReadOnly(),
                Array.Empty<PdfRedactionMatch>(), findings, DescribeCriteria(search), PdfRedactionPlan.ComputeSourceSha256(pdf),
                PdfRedactionPlan.CapturePageIdentities(readDocument, areas));
        }
        PdfRedactionPlan planned = Plan(pdf, areas, layoutOptions, readOptions, search.CancellationToken);
        search.CancellationToken.ThrowIfCancellationRequested();
        return new PdfRedactionPlan(planned.Preflight, planned.Areas, planned.Matches,
            planned.Findings.Concat(selectionFindings).ToArray(), DescribeCriteria(search), planned.SourceSha256,
            planned.PageIdentities, planned.ReviewedTextObjectScopes, search.MatchCase, search.RegexOptions, search.RegexTimeout);
    }
}
