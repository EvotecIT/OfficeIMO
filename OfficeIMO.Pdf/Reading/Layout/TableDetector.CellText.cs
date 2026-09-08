namespace OfficeIMO.Pdf;

internal static partial class TableDetector {
    internal static string[] SplitBySplits(TextLayoutEngine.TextLine line, List<double> splits) {
        var fragments = new List<PdfTextSpan>[splits.Count + 1];
        for (int index = 0; index < fragments.Length; index++) fragments[index] = new List<PdfTextSpan>();
        foreach (PdfTextSpan span in line.Spans) {
            int column = 0;
            while (column < splits.Count && span.X >= splits[column]) column++;
            fragments[column].Add(span);
        }
        return fragments.Select(ComposeCell).ToArray();
    }

    private static string[] SplitByGaps(TextLayoutEngine.TextLine line) {
        var cells = new List<string>();
        var current = new List<PdfTextSpan>();
        foreach (PdfTextSpan span in line.Spans) {
            if (current.Count > 0) {
                PdfTextSpan previous = current[current.Count - 1];
                double gap = span.X - (previous.X + Math.Max(0D, previous.Advance));
                if (gap > Math.Max(18D, Math.Max(previous.FontSize, span.FontSize) * 2D)) {
                    cells.Add(ComposeCell(current));
                    current = new List<PdfTextSpan>();
                }
            }
            current.Add(span);
        }
        if (current.Count > 0) cells.Add(ComposeCell(current));
        return cells.ToArray();
    }

    private static string ComposeCell(List<PdfTextSpan> spans) => spans.Count == 0
        ? string.Empty
        : TextLayoutEngine.BuildLine(spans, null).Text.Trim();

    private static bool HasExplicitBoundarySpace(PdfTextSpan previous, PdfTextSpan current) =>
        previous.LogicalTrailingSpace || current.LogicalLeadingSpace ||
        (previous.Text.Length > 0 && char.IsWhiteSpace(previous.Text[previous.Text.Length - 1])) ||
        (current.Text.Length > 0 && char.IsWhiteSpace(current.Text[0]));
}
