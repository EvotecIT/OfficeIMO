using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Distributes each justified line's free space over its word separators and returns the space
    /// added per line. The last line of the block and lines ending in a forced break follow
    /// text-align-last; with its auto value they align to start.
    /// </summary>
    private static double[] ApplyInlineLineJustification(IReadOnlyList<InlineLine> lines, double width,
        HtmlRenderBoxStyle paragraphStyle, out JustificationLimit limits) {
        var added = new double[lines.Count];
        limits = JustificationLimit.None;
        // Right-to-left paragraphs whose lines cannot be bidi-resolved, because they hold
        // atomic boxes or markers, place segments right to left in logical order. Splitting a
        // left-to-right phrase at its separators would reverse its words, so they stay at start.
        bool unresolvedRightToLeft = string.Equals(paragraphStyle.Direction, "rtl", StringComparison.Ordinal)
            && lines.SelectMany(static line => line.Segments).Any(static segment =>
                segment.Run.AtomicBlock != null || segment.Run.IsEmptyInlineBox
                || segment.Run.RunningStringElement != null || segment.Run.PositionedMarkerElement != null);
        for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
            InlineLine line = lines[lineIndex];
            ClearInlineLineJustification(line);
            if (!JustifiesInlineLine(lines, lineIndex, paragraphStyle)) continue;
            if (unresolvedRightToLeft) {
                limits |= JustificationLimit.RightToLeftWithInlineBoxes;
                continue;
            }
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
            if (opportunities == 0) {
                // CJK text has no word separators; browsers justify it between characters.
                if (line.Segments.Any(static segment => ContainsCjkText(segment.Text))) limits |= JustificationLimit.CjkWithoutSeparators;
                continue;
            }

            double perOpportunity = extra / opportunities;
            for (int index = first + 1; index < last; index++) {
                InlineSegment segment = line.Segments[index];
                int count = CountJustificationOpportunities(segment);
                if (count > 0) {
                    segment.JustificationAdvance = perOpportunity * count;
                    if (segment.Run.Style.TextJustifyInterCharacter) limits |= JustificationLimit.InterCharacter;
                }
            }
            added[lineIndex] = extra;
        }
        return added;
    }

    [Flags]
    private enum JustificationLimit {
        None = 0,
        RightToLeftWithInlineBoxes = 1,
        InterCharacter = 4,
        CjkWithoutSeparators = 2
    }

    private static bool ContainsCjkText(string text) {
        foreach (char value in text) {
            if (value is >= '\u3040' and <= '\u30FF' or >= '\u3400' and <= '\u4DBF' or >= '\u4E00' and <= '\u9FFF'
                or >= '\uF900' and <= '\uFAFF') return true;
        }
        return false;
    }

    // Justified lines that are approximated are reported once per paragraph and case.
    private void ReportJustificationLimits(double[] added, JustificationLimit limits, HtmlRenderBoxStyle paragraphStyle,
        AngleSharp.Dom.IElement? formattingContainer) {
        string source = formattingContainer != null ? HtmlRenderStyleResolver.DescribeSource(formattingContainer) : paragraphStyle.SemanticRole;
        if ((limits & JustificationLimit.InterCharacter) != 0)
            ReportJustificationLimit(formattingContainer, source, "text-justify=inter-character",
                "Inter-character justification is not implemented; justified text was spaced between words.");
        if ((limits & JustificationLimit.RightToLeftWithInlineBoxes) != 0)
            ReportJustificationLimit(formattingContainer, source, "text-align=justify;rtl-inline-boxes",
                "Right-to-left lines with inline boxes were not justified; they were aligned to start.");
        if ((limits & JustificationLimit.CjkWithoutSeparators) != 0)
            ReportJustificationLimit(formattingContainer, source, "text-align=justify;cjk",
                "Lines of CJK text without word separators were not justified; they were aligned to start.");
    }

    private void ReportJustificationLimit(object? owner, string source, string detail, string message) {
        if (!_reportedJustificationLimits.Add((owner ?? source, detail))) return;
        _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.CssPropertyUnsupported, message,
            HtmlDiagnosticSeverity.Warning, source, detail, OfficeConversionLossKind.Approximation);
    }

    private readonly HashSet<(object Owner, string Detail)> _reportedJustificationLimits = new();

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
        // text-justify applies to the text it is set on, so a none span keeps its spaces.
        if (segment.Run.Style.TextJustifyNone) return 0;
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
