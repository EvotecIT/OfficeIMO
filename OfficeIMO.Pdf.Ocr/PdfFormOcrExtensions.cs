using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Ocr;

/// <summary>Recognition of visible values inside existing PDF form widgets.</summary>
public static class PdfFormOcrExtensions {
    /// <summary>Recognizes a captured source and proposes values inside existing widget bounds without mutation.</summary>
    /// <remarks>Uses original-page geometry, including crop/rotation/UserUnit and scan transforms. Native-overlap words
    /// remain useful evidence for filling and do not become a searchable layer. Hidden widgets are excluded. Arbitrary
    /// form actions are not executed; only the canonical bounded inert numeric helper profile is assessed; other scripts are rejected.
    /// Repeated selected physical pages are recognized once. Completion waits for started provider calls and cancellation callbacks
    /// to settle; a provider that ignores cancellation can delay completion indefinitely.</remarks>
    public static async Task<PdfFormOcrReview> PrepareFormOcrAsync(this PdfDocument document, IOcrEngine engine,
        PdfOcrMergeOptions? options = null, CancellationToken cancellationToken = default) {
        Guard.NotNull(document, nameof(document)); Guard.NotNull(engine, nameof(engine));
        cancellationToken.ThrowIfCancellationRequested();
        var snapshot = PdfDocument.Load(document.ToBytes(cancellationToken), document.ReadOptions);
        var effective = options?.Clone() ?? new PdfOcrMergeOptions();
        effective.AwaitProviderSettlement = true;
        var info = snapshot.Inspect(snapshot.ReadOptions, cancellationToken);
        PdfOcrExtensions.NormalizeReviewPageSelection(effective, info.PageCount);
        var fields = info.FormFields;
        if (fields.Count > 5000) throw new InvalidOperationException("Form recognition supports at most 5000 existing fields per review.");
        var layouts = snapshot.GetPageLayouts(snapshot.ReadOptions, cancellationToken).ToDictionary(page => page.PageNumber);
        var ocr = await snapshot.ReadWithOcrAsync(engine, effective, cancellationToken).ConfigureAwait(false);
        var pages = ocr.Pages.ToDictionary(page => page.PageNumber);
        var proposals = new List<PdfFormOcrProposal>();
        var assignments = new Dictionary<PdfOcrWordEvidence, HashSet<string>>();
        long comparisons = 0;
        foreach (var field in fields.Where(field => !string.IsNullOrWhiteSpace(field.Name))) {
            var evidence = new List<PdfFormOcrEvidence>();
            var candidates = new List<string>();
            foreach (var widget in field.Widgets) {
                cancellationToken.ThrowIfCancellationRequested();
                if (widget.IsHidden || widget.IsInvisible || widget.IsNoView || widget.PageNumber is not int number ||
                    !pages.TryGetValue(number, out var page) || widget.Width <= 0 || widget.Height <= 0) continue;
                var bounds = layouts[number].MapUserSpaceRectangleToVisual(widget.X1, widget.Y1, widget.X2, widget.Y2);
                var words = new List<PdfOcrWordEvidence>();
                foreach (var word in page.WordEvidence) {
                    if (++comparisons > 5_000_000) throw new InvalidOperationException("Form recognition geometry comparison budget exceeded.");
                    cancellationToken.ThrowIfCancellationRequested();
                    var quad = word.Word.Geometry;
                    double area = Math.Max(0, quad.Right - quad.Left) * Math.Max(0, quad.Bottom - quad.Top);
                    double intersection = Math.Max(0, Math.Min(bounds.Right, quad.Right) - Math.Max(bounds.Left, quad.Left)) *
                        Math.Max(0, Math.Min(bounds.Bottom, quad.Bottom) - Math.Max(bounds.Top, quad.Top));
                    if (area <= 0 || intersection / area < 0.5) continue;
                    words.Add(word); evidence.Add(new(number, bounds, word));
                    if (!assignments.TryGetValue(word, out var names)) assignments.Add(word, names = new(StringComparer.Ordinal));
                    names.Add(field.Name!);
                }
                if (words.Count > 0) candidates.Add(string.Join(" ", PdfOcrLogicalDocumentBuilder.OrderWordsForLogicalReading(words.Select(item => item.Word).ToArray(),
                    ocr.Document.Pages.Single(item => item.PageNumber == number), effective.ReadOptions.LayoutOptions.ReadingDirection, cancellationToken).Select(word => word.Text)));
            }
            if (evidence.Count == 0) continue;
            string? rejection = field.IsReadOnly ? "This field is read-only." :
                !field.IsTextField && !field.IsChoiceField ? "Use the form inspector for buttons, signatures and unsupported fields." :
                PdfFormFieldValueAssessment.Assess(field, PdfFormFieldValue.From(string.Empty)).Issues.Any(issue => issue.Code == PdfFormFieldValueIssueCode.UnsupportedScriptConstraint) ?
                "This field has script constraints, including possible numeric rules, which OCR review does not execute or validate. Fill it manually in a qualified form application." :
                field.Widgets.Any(widget => widget.IsReadOnly) ? "A widget for this field is read-only." : null;
            proposals.Add(new(field, candidates.FirstOrDefault() ?? string.Empty, evidence.AsReadOnly(),
                candidates.Distinct(StringComparer.Ordinal).Count() > 1, rejection));
        }
        // Overlap is explicit evidence, never an automatic choice of destination.
        proposals = proposals.Select(proposal => new PdfFormOcrProposal(proposal.Field, proposal.SuggestedValue, proposal.Evidence,
            proposal.IsAmbiguous || proposal.Evidence.Any(item => assignments[item.Word].Count > 1), proposal.RejectionReason)).ToList();
        cancellationToken.ThrowIfCancellationRequested();
        return new PdfFormOcrReview(snapshot, effective, ocr, proposals.AsReadOnly());
    }
}
