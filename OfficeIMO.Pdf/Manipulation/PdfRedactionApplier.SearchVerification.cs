using System.Text.RegularExpressions;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionApplier {
    // Search rectangles depend on source text geometry. Recheck every requested
    // criterion on affected pages after rewriting, even when several criteria
    // matched the same source block and only the first one labelled its area.
    internal static void VerifySearchedTextRemoved(byte[] output, byte[] source, PdfRedactionPlan plan,
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
        var workBudget = new PdfRedactionSearchWorkBudget("search verification");
        foreach (KeyValuePair<int, PdfRedactionArea[]> page in areasByPage) {
            int pageNumber = page.Key;
            cancellationToken.ThrowIfCancellationRequested();
            if (pageNumber < 1 || pageNumber > rewritten.Pages.Count)
                throw new InvalidOperationException("The rewritten PDF is missing a searched redaction page.");
            string remaining = rewritten.Pages[pageNumber - 1].ExtractText(cancellationToken);
            foreach (string literal in literals) {
                cancellationToken.ThrowIfCancellationRequested();
                workBudget.ChargeTextScan(remaining, literal);
                if (PdfTextSearchNormalization.ContainsExact(remaining, literal, comparison)) ThrowSurvivingText(pageNumber);
                // ExtractText does not preserve every cell wrap or embedded /ActualText line break.
                // Reuse the planner's native normalization so a surviving wrapped occurrence fails closed.
                var nativeOptions = new PdfTextSearchOptions {
                    MatchCase = plan.SearchMatchCase,
                    IncludeTextRenderingMode3 = true,
                    PageNumbers = new[] { pageNumber }
                };
                if (PdfTextEditor.Find(output, literal, nativeOptions, outputOptions, workBudget).Count > 0)
                    ThrowSurvivingText(pageNumber);
            }
            PdfLogicalTextBlock[] rewrittenBlocks = rewrittenBlocksByPage != null &&
                rewrittenBlocksByPage.TryGetValue(pageNumber, out PdfLogicalTextBlock[]? pageRewrittenBlocks)
                ? pageRewrittenBlocks : Array.Empty<PdfLogicalTextBlock>();
            if (rewrittenLogical != null) {
                foreach (PdfLogicalTextBlock block in rewrittenBlocks) {
                    cancellationToken.ThrowIfCancellationRequested();
                    foreach (Regex regex in regexes) if (workBudget.IsMatch(regex, block.Text)) ThrowSurvivingText(pageNumber);
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
                Dictionary<int, string> wrappedLiterals = PdfRedactionPlanner.MatchLiteralsAcrossBlocks(
                    sourceBlocks, literals, comparison, workBudget, static _ => true, cancellationToken,
                    readOptions?.Limits.MaxTextSearchFlowComparisons ?? PdfReadLimits.DefaultMaxTextSearchFlowComparisons);
                for (int blockIndex = 0; blockIndex < sourceBlocks.Length; blockIndex++) {
                    PdfLogicalTextBlock block = sourceBlocks[blockIndex];
                    cancellationToken.ThrowIfCancellationRequested();
                    bool selected = selectedKinds.Contains(block.Kind) || wrappedLiterals.ContainsKey(blockIndex);
                    for (int i = 0; !selected && i < literals.Length; i++) {
                        workBudget.ChargeTextScan(block.Text, literals[i]);
                        selected = Contains(block.Text, literals[i], comparison);
                    }
                    for (int i = 0; !selected && i < regexes.Length; i++)
                        selected = workBudget.IsMatch(regexes[i], block.Text);
                    if (!selected) continue;
                    foreach (PdfTextSpan span in block.Spans) {
                        if (string.IsNullOrEmpty(span.Text)) continue;
                        PdfTextSpanBounds sourceBounds = PdfTextSpanGeometry.GetAxisAlignedBounds(span);
                        foreach (PdfRedactionArea area in page.Value) {
                            cancellationToken.ThrowIfCancellationRequested();
                            workBudget.Charge(1L);
                            if (!area.IntersectsRectangle(sourceBounds.Left, sourceBounds.Bottom, sourceBounds.Width, sourceBounds.Height))
                                continue;
                            foreach ((PdfTextSpan candidateSpan, PdfTextSpanBounds candidateBounds) in rewrittenSpans) {
                                cancellationToken.ThrowIfCancellationRequested();
                                workBudget.Charge(1L);
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

    // Select source blocks with the planner's line-break and hyphenation-tolerant matching so wrapped occurrences stay covered.
    private static bool Contains(string text, string value, StringComparison comparison) =>
        PdfTextSearchNormalization.Contains(text, value, comparison);

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
