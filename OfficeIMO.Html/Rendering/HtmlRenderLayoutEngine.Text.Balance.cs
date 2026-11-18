namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    // Balancing lays a block out several times, so it is limited to short blocks, as in browsers.
    private const int MaximumBalancedLines = 6;
    private const int MaximumBalancedCharacters = 1000;
    private const int BalanceIterations = 12;

    private static bool CanBalanceInlineLines(IReadOnlyList<HtmlInlineRun> runs, HtmlRenderBoxStyle paragraphStyle,
        int skipLogicalCharacters, int? maximumLines, InlinePaintCapture? paintCapture) {
        if (paragraphStyle.TextWrapStyle != "balance" || paragraphStyle.PreventTextWrapping
            || skipLogicalCharacters > 0 || maximumLines.HasValue || paintCapture != null
            || paragraphStyle.LineClamp.HasValue) return false;
        int characters = 0;
        foreach (HtmlInlineRun run in runs) {
            // Inline box edges record their first and last content while lines are admitted.
            if (run.InlineEdgeScopes.Count > 0 || run.PositionedMarkerElement != null
                || run.RunningStringElement != null || run.RunningElementAssignment != null) return false;
            characters += run.Text.Length;
            if (characters > MaximumBalancedCharacters) return false;
        }
        return true;
    }

    /// <summary>
    /// text-wrap: balance. Finds the narrowest line length that keeps the block's line count, so its
    /// lines have similar lengths, and paints the lines in the full width for alignment.
    /// </summary>
    private HtmlInlineLayout LayoutBalancedInlineRuns(IReadOnlyList<HtmlInlineRun> runs, double width,
        HtmlRenderBoxStyle paragraphStyle, AngleSharp.Dom.IElement? formattingContainer) {
        int CountLines(double candidate, bool edgesPrepared) {
            int count = 0;
            LayoutInlineRuns(runs, candidate, paragraphStyle, formattingContainer, inlineEdgesPrepared: edgesPrepared,
                renderWidth: width, admittedLines: value => count = value);
            return count;
        }

        int target = CountLines(width, edgesPrepared: false);
        double balanced = width;
        if (target >= 2 && target <= MaximumBalancedLines) {
            double low = width / target, high = width;
            for (int iteration = 0; iteration < BalanceIterations && high - low > 0.5D; iteration++) {
                CheckCancellation();
                double middle = (low + high) / 2D;
                if (CountLines(middle, edgesPrepared: true) <= target) high = middle;
                else low = middle;
            }
            balanced = high;
        }
        return LayoutInlineRuns(runs, balanced, paragraphStyle, formattingContainer, inlineEdgesPrepared: true, renderWidth: width);
    }
}
