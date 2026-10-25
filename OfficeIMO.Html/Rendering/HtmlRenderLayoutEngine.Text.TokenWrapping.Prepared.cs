using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Wraps generated provider text by validated source clusters without reshaping their glyphs.</summary>
    private bool TryAddPreparedBrokenToken(ICollection<InlineLine> lines, ref InlineLine line, HtmlInlineRun run,
        string token, string logicalToken, double width, int logicalStartProgress, bool completesToken) {
        var segment = new InlineSegment(token, MeasureInlineText(token, run.Style, run), run, logicalToken);
        // Reuse the same complete/nonnegative/additive partition as natural bidi.
        // It proves source-grapheme boundaries and retains the accepted context.
        if (!TryPartitionPreparedBidiWidths(segment, OfficeTextElements.Split(token),
                OfficeTextElements.Split(logicalToken), out _) || segment.PreparedShaping == null) {
            if (TryMeasureWithConfiguredProvider(logicalToken, run.Style, out _))
                ReportProviderContextLoss(run, "generated-wrap-cluster-context-not-retained");
            return false;
        }
        OfficePreparedTextShaping context = segment.PreparedShaping;
        var starts = new SortedSet<int>();
        for (int index = 0; index < context.Result.Glyphs.Count; index++) {
            CheckCancellation();
            starts.Add(context.Result.GetLogicalClusterStart(index));
        }
        int[] boundaries = starts.Concat(new[] { context.Text.Length }).ToArray();
        for (int index = 0; index + 1 < boundaries.Length; index++) {
            CheckCancellation();
            int start = boundaries[index], length = boundaries[index + 1] - start;
            // Partitioning above validated every same complete slice. Neither
            // the immutable source nor the selected font changes between them.
            if (!context.TrySlice(start, length, out OfficePreparedTextShaping? slice) || slice == null)
                throw new InvalidOperationException("A validated generated-text cluster could not retain its glyph slice.");
            double advance = slice.Measure(GetEffectiveTextFont(run.Style).Size);
            bool finalPiece = completesToken && index + 2 == boundaries.Length;
            if (line.HasFlowContent && line.PreviewAdvance(run, advance, finalPiece) > line.ResolveAvailableWidth(width) + 0.0001D) {
                TrimTrailingWhitespace(line);
                lines.Add(line);
                line = new InlineLine();
            }
            line.Add(new InlineSegment(slice.Text, advance, run, slice.Text,
                logicalEndProgress: logicalStartProgress + start + length, preparedShaping: slice));
        }
        return true;
    }
}
