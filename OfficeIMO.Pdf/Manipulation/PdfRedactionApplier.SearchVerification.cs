using System.Text.RegularExpressions;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionApplier {
    // Search rectangles depend on source text geometry. Recheck every requested
    // criterion on affected pages after rewriting, even when several criteria
    // matched the same source block and only the first one labelled its area.
    private static void VerifySearchedTextRemoved(byte[] output, byte[] source, PdfRedactionPlan plan,
        PdfTextLayoutOptions? layoutOptions, PdfLoadOptions? readOptions,
        PdfGeneratedOutputGrowth generatedGrowth, CancellationToken cancellationToken) {
        if (plan.SearchCriteria.Count == 0 || plan.Areas.Count == 0) return;
        int[] affectedPages = plan.Areas.Select(static area => area.PageNumber).Distinct().ToArray();
        string[] literals = plan.SearchCriteria.Where(static criterion => criterion.StartsWith("literal:", StringComparison.Ordinal))
            .Select(static criterion => criterion.Substring("literal:".Length)).ToArray();
        Regex[] regexes = plan.SearchCriteria.Where(static criterion => criterion.StartsWith("regex:", StringComparison.Ordinal))
            .Select(criterion => new Regex(criterion.Substring("regex:".Length), plan.SearchRegexOptions, plan.SearchRegexTimeout)).ToArray();
        PdfLogicalElementKind[] kinds = plan.SearchCriteria.Where(static criterion => criterion.StartsWith("logical-kind:", StringComparison.Ordinal))
            .Select(static criterion => ParseLogicalKind(criterion)).ToArray();
        if (literals.Length == 0 && regexes.Length == 0 && kinds.Length == 0) return;

        PdfLoadOptions outputOptions = PdfLoadOptions.ForGeneratedOutput(readOptions, source, output, generatedGrowth);
        PdfReadDocument rewritten = PdfReadDocument.Open(output, outputOptions, cancellationToken);
        PdfDocumentReadResult? rewrittenLogical = regexes.Length > 0 || kinds.Length > 0
            ? PdfDocumentReadResult.From(rewritten, layoutOptions) : null;
        PdfDocumentReadResult? sourceLogical = kinds.Length > 0
            ? PdfDocumentReadResult.From(PdfReadDocument.Open(source, readOptions, cancellationToken), layoutOptions) : null;
        StringComparison comparison = plan.SearchMatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        foreach (int pageNumber in affectedPages) {
            cancellationToken.ThrowIfCancellationRequested();
            if (pageNumber < 1 || pageNumber > rewritten.Pages.Count)
                throw new InvalidOperationException("The rewritten PDF is missing a searched redaction page.");
            string remaining = rewritten.Pages[pageNumber - 1].ExtractText(cancellationToken);
            foreach (string literal in literals) {
                cancellationToken.ThrowIfCancellationRequested();
                if (Contains(remaining, literal, comparison)) ThrowSurvivingText(pageNumber);
            }
            if (rewrittenLogical != null) {
                foreach (PdfLogicalTextBlock block in rewrittenLogical.TextBlocks) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (block.PageNumber != pageNumber) continue;
                    foreach (Regex regex in regexes) if (regex.IsMatch(block.Text)) ThrowSurvivingText(pageNumber);
                }
            }
            if (sourceLogical != null) {
                foreach (PdfLogicalTextBlock block in sourceLogical.TextBlocks) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (block.PageNumber != pageNumber || !kinds.Contains(block.Kind)) continue;
                    foreach (PdfTextSpan span in block.Spans) {
                        if (string.IsNullOrEmpty(span.Text)) continue;
                        PdfTextSpanBounds sourceBounds = PdfTextSpanGeometry.GetAxisAlignedBounds(span);
                        bool selected = plan.Areas.Any(area => area.PageNumber == pageNumber &&
                            area.IntersectsRectangle(sourceBounds.Left, sourceBounds.Bottom, sourceBounds.Width, sourceBounds.Height));
                        if (!selected) continue;
                        foreach (PdfLogicalTextBlock rewrittenBlock in rewrittenLogical!.TextBlocks) {
                            cancellationToken.ThrowIfCancellationRequested();
                            if (rewrittenBlock.PageNumber != pageNumber) continue;
                            foreach (PdfTextSpan rewrittenSpan in rewrittenBlock.Spans) {
                                if (!Contains(rewrittenSpan.Text, span.Text, StringComparison.Ordinal)) continue;
                                PdfTextSpanBounds rewrittenBounds = PdfTextSpanGeometry.GetAxisAlignedBounds(rewrittenSpan);
                                if (plan.Areas.Any(area => area.PageNumber == pageNumber &&
                                    area.IntersectsRectangle(sourceBounds.Left, sourceBounds.Bottom, sourceBounds.Width, sourceBounds.Height) &&
                                    area.IntersectsRectangle(rewrittenBounds.Left, rewrittenBounds.Bottom, rewrittenBounds.Width, rewrittenBounds.Height)))
                                    ThrowSurvivingText(pageNumber);
                            }
                        }
                    }
                }
            }
        }
    }

    private static bool Contains(string text, string value, StringComparison comparison) {
#if NET6_0_OR_GREATER
        return text.Contains(value, comparison);
#else
        return text.IndexOf(value, comparison) >= 0;
#endif
    }

    private static PdfLogicalElementKind ParseLogicalKind(string criterion) {
#if NET6_0_OR_GREATER
        return Enum.Parse<PdfLogicalElementKind>(criterion.AsSpan("logical-kind:".Length));
#else
        return (PdfLogicalElementKind)Enum.Parse(typeof(PdfLogicalElementKind), criterion.Substring("logical-kind:".Length));
#endif
    }

    private static void ThrowSurvivingText(int pageNumber) =>
        throw new InvalidOperationException("The rewritten PDF still contains searched text on page " +
            pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".");
}
