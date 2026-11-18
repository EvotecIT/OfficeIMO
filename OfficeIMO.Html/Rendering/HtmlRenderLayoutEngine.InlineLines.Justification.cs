using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Distributes each justified line's free space over its word separators and returns the space
    /// added per line. The last line of the block and lines ending in a forced break follow
    /// text-align-last; with its auto value they align to start.
    /// </summary>
    private static double[] ApplyInlineLineJustification(IReadOnlyList<InlineLine> lines, double width,
        HtmlRenderBoxStyle paragraphStyle) {
        var added = new double[lines.Count];
        for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
            InlineLine line = lines[lineIndex];
            ClearInlineLineJustification(line);
            if (!JustifiesInlineLine(lines, lineIndex, paragraphStyle)) continue;
            double extra = line.ResolveAvailableWidth(width) - line.Width;
            if (extra <= 0.01D) continue;

            int first = -1, last = -1;
            for (int index = 0; index < line.Segments.Count; index++) {
                if (!IsJustificationContent(line.Segments[index])) continue;
                if (first < 0) first = index;
                last = index;
            }
            int opportunities = 0;
            for (int index = first + 1; index < last; index++) opportunities += CountJustificationOpportunities(line.Segments[index]);
            if (opportunities == 0) continue;

            double perOpportunity = extra / opportunities;
            for (int index = first + 1; index < last; index++) {
                InlineSegment segment = line.Segments[index];
                int count = CountJustificationOpportunities(segment);
                if (count > 0) segment.JustificationAdvance = perOpportunity * count;
            }
            added[lineIndex] = extra;
        }
        return added;
    }

    private static void ClearInlineLineJustification(InlineLine line) {
        foreach (InlineSegment segment in line.Segments) segment.JustificationAdvance = 0D;
    }

    /// <summary>The alignment of a line that is not justified, or the fallback when it has no word separator.</summary>
    private static OfficeTextAlignment ResolveInlineLineAlignment(IReadOnlyList<InlineLine> lines, int lineIndex,
        HtmlRenderBoxStyle paragraphStyle) {
        if (IsLastInlineLine(lines, lineIndex) && paragraphStyle.TextAlignLast is OfficeTextAlignment last) {
            return last != OfficeTextAlignment.Justify ? last
                : string.Equals(paragraphStyle.Direction, "rtl", StringComparison.Ordinal)
                    ? OfficeTextAlignment.Right : OfficeTextAlignment.Left;
        }
        // Justified text stores its start alignment for lines that are not justified.
        return paragraphStyle.Alignment;
    }

    private static bool JustifiesInlineLine(IReadOnlyList<InlineLine> lines, int lineIndex, HtmlRenderBoxStyle paragraphStyle) =>
        IsLastInlineLine(lines, lineIndex)
            ? paragraphStyle.TextAlignLast == OfficeTextAlignment.Justify
            : paragraphStyle.JustifyText;

    private static bool IsLastInlineLine(IReadOnlyList<InlineLine> lines, int lineIndex) =>
        lineIndex == lines.Count - 1 || lines[lineIndex].HasForcedBreak;

    private static bool IsJustificationContent(InlineSegment segment) =>
        InlineLine.ContributesFlowContent(segment.Run)
        && (segment.Run.AtomicBlock != null || segment.Text.Length > 0 && !IsWhitespaceToken(segment.Text));

    // Word separators between content are the justification opportunities. Collapsed
    // white space is one separator; preserved spaces each count, while tabs keep their stops.
    private static int CountJustificationOpportunities(InlineSegment segment) {
        if (segment.Run.AtomicBlock != null || !InlineLine.ContributesFlowContent(segment.Run)
            || segment.Text.Length == 0 || !IsWhitespaceToken(segment.Text)) return 0;
        if (!segment.Run.Style.PreserveWhitespace) return 1;
        int count = 0;
        foreach (char value in segment.Text) {
            if (value == ' ' || value == ' ') count++;
        }
        return count;
    }

    /// <summary>Paints a justified separator across its whole advance so decorations stay continuous.</summary>
    private static IReadOnlyList<InlinePaintSegment> WidenJustifiedSeparator(IReadOnlyList<InlinePaintSegment> segments,
        double justification) {
        if (justification <= 0D || segments.Count == 0) return segments;
        var widened = new InlinePaintSegment[segments.Count];
        for (int index = 0; index < segments.Count; index++) widened[index] = segments[index];
        InlinePaintSegment last = widened[widened.Length - 1];
        widened[widened.Length - 1] = new InlinePaintSegment(last.Text, last.X, last.Width + justification,
            last.Advance + justification, last.LogicalOrder, last.BidiResolved, last.PreparedShaping);
        return widened;
    }
}
