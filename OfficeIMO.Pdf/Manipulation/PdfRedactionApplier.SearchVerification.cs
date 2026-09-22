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
        Dictionary<int, PdfRedactionArea[]> areasByPage = plan.Areas.GroupBy(static area => area.PageNumber)
            .ToDictionary(static group => group.Key, static group => group.ToArray());
        string[] literals = plan.SearchCriteria.Where(static criterion => criterion.StartsWith("literal:", StringComparison.Ordinal))
            .Select(static criterion => criterion.Substring("literal:".Length)).ToArray();
        Regex[] regexes = plan.SearchCriteria.Where(static criterion => criterion.StartsWith("regex:", StringComparison.Ordinal))
            .Select(criterion => new Regex(criterion.Substring("regex:".Length), plan.SearchRegexOptions, plan.SearchRegexTimeout)).ToArray();
        PdfLogicalElementKind[] kinds = plan.SearchCriteria.Where(static criterion => criterion.StartsWith("logical-kind:", StringComparison.Ordinal))
            .Select(static criterion => ParseLogicalKind(criterion)).ToArray();
        if (literals.Length == 0 && regexes.Length == 0 && kinds.Length == 0) return;

        PdfLoadOptions outputOptions = PdfLoadOptions.ForGeneratedOutput(readOptions, source, output, generatedGrowth);
        PdfReadDocument rewritten = PdfReadDocument.Open(output, outputOptions, cancellationToken);
        PdfDocumentReadResult? rewrittenLogical = literals.Length > 0 || regexes.Length > 0 || kinds.Length > 0
            ? PdfDocumentReadResult.From(rewritten, layoutOptions, cancellationToken) : null;
        PdfDocumentReadResult? sourceLogical = literals.Length > 0 || regexes.Length > 0 || kinds.Length > 0
            ? PdfDocumentReadResult.From(PdfReadDocument.Open(source, readOptions, cancellationToken), layoutOptions, cancellationToken) : null;
        Dictionary<int, PdfLogicalTextBlock[]>? rewrittenBlocksByPage = rewrittenLogical?.TextBlocks
            .GroupBy(static block => block.PageNumber)
            .ToDictionary(static group => group.Key, static group => group.ToArray());
        Dictionary<int, PdfLogicalTextBlock[]>? sourceBlocksByPage = sourceLogical?.TextBlocks
            .GroupBy(static block => block.PageNumber)
            .ToDictionary(static group => group.Key, static group => group.ToArray());
        var selectedKinds = new HashSet<PdfLogicalElementKind>(kinds);
        StringComparison comparison = plan.SearchMatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        long remainingLogicalVerificationWork = 20_000_000L;
        foreach (KeyValuePair<int, PdfRedactionArea[]> page in areasByPage) {
            int pageNumber = page.Key;
            cancellationToken.ThrowIfCancellationRequested();
            if (pageNumber < 1 || pageNumber > rewritten.Pages.Count)
                throw new InvalidOperationException("The rewritten PDF is missing a searched redaction page.");
            string remaining = rewritten.Pages[pageNumber - 1].ExtractText(cancellationToken);
            foreach (string literal in literals) {
                cancellationToken.ThrowIfCancellationRequested();
                if (Contains(remaining, literal, comparison)) ThrowSurvivingText(pageNumber);
            }
            PdfLogicalTextBlock[] rewrittenBlocks = rewrittenBlocksByPage != null &&
                rewrittenBlocksByPage.TryGetValue(pageNumber, out PdfLogicalTextBlock[]? pageRewrittenBlocks)
                ? pageRewrittenBlocks : Array.Empty<PdfLogicalTextBlock>();
            if (rewrittenLogical != null) {
                foreach (PdfLogicalTextBlock block in rewrittenBlocks) {
                    cancellationToken.ThrowIfCancellationRequested();
                    foreach (Regex regex in regexes) if (regex.IsMatch(block.Text)) ThrowSurvivingText(pageNumber);
                }
            }
            if (sourceLogical != null) {
                PdfLogicalTextBlock[] sourceBlocks = sourceBlocksByPage != null &&
                    sourceBlocksByPage.TryGetValue(pageNumber, out PdfLogicalTextBlock[]? pageSourceBlocks)
                    ? pageSourceBlocks : Array.Empty<PdfLogicalTextBlock>();
                var rewrittenSpans = new List<(PdfTextSpan Span, PdfTextSpanBounds Bounds)>();
                foreach (PdfLogicalTextBlock block in rewrittenBlocks) {
                    cancellationToken.ThrowIfCancellationRequested();
                    foreach (PdfTextSpan span in block.Spans)
                        rewrittenSpans.Add((span, PdfTextSpanGeometry.GetAxisAlignedBounds(span)));
                }
                foreach (PdfLogicalTextBlock block in sourceBlocks) {
                    cancellationToken.ThrowIfCancellationRequested();
                    bool selected = selectedKinds.Contains(block.Kind) ||
                        literals.Any(literal => Contains(block.Text, literal, comparison)) ||
                        regexes.Any(regex => regex.IsMatch(block.Text));
                    if (!selected) continue;
                    foreach (PdfTextSpan span in block.Spans) {
                        if (string.IsNullOrEmpty(span.Text)) continue;
                        PdfTextSpanBounds sourceBounds = PdfTextSpanGeometry.GetAxisAlignedBounds(span);
                        foreach (PdfRedactionArea area in page.Value) {
                            cancellationToken.ThrowIfCancellationRequested();
                            ConsumeLogicalVerificationWork(ref remainingLogicalVerificationWork, 1L);
                            if (!area.IntersectsRectangle(sourceBounds.Left, sourceBounds.Bottom, sourceBounds.Width, sourceBounds.Height))
                                continue;
                            foreach ((PdfTextSpan candidateSpan, PdfTextSpanBounds candidateBounds) in rewrittenSpans) {
                                cancellationToken.ThrowIfCancellationRequested();
                                ConsumeLogicalVerificationWork(ref remainingLogicalVerificationWork, 1L);
                                if (!string.IsNullOrWhiteSpace(candidateSpan.Text) &&
                                    sourceBounds.Left < candidateBounds.Right && sourceBounds.Right > candidateBounds.Left &&
                                    sourceBounds.Bottom < candidateBounds.Top && sourceBounds.Top > candidateBounds.Bottom)
                                    ThrowSurvivingText(pageNumber);
                            }
                        }
                    }
                }
            }
        }
    }

    private static void ConsumeLogicalVerificationWork(ref long remaining, long work) {
        if (work > remaining)
            throw new InvalidDataException("PDF redaction search verification exceeds its logical text work limit.");
        remaining -= work;
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
