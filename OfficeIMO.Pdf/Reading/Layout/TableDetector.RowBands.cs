namespace OfficeIMO.Pdf;

internal static partial class TableDetector {
    /// <summary>
    /// Recognizes vertically centred wrapped cells from leading and trailing text
    /// in the same column around an anchor. A single-line first row does not
    /// establish vertical alignment for the remaining rows.
    /// </summary>
    private static HashSet<int> FindCenteredWrappedAnchors(
        List<(TextLayoutEngine.TextLine Line, PositionedRow Row)> anchors,
        IReadOnlyList<TextLayoutEngine.TextLine> lines, double[] splits,
        int firstBodyIndex, double headerBoundary, double rowPitch, double left, double right,
        Action<long>? consumeWork, Action? cancellationCheck) {
        var centered = new HashSet<int>();
        for (int index = firstBodyIndex; index < anchors.Count; index++) {
            cancellationCheck?.Invoke();
            double y = anchors[index].Row.Y;
            double upper = index == firstBodyIndex ? headerBoundary : (anchors[index - 1].Row.Y + y) / 2D;
            double lower = index + 1 < anchors.Count ? (y + anchors[index + 1].Row.Y) / 2D : y - rowPitch / 2D;
            double fontSize = anchors[index].Line.Spans.Select(static span => span.FontSize).DefaultIfEmpty(1D).Max();
            var above = Enumerable.Repeat(double.PositiveInfinity, splits.Length + 1).ToArray();
            var below = Enumerable.Repeat(double.PositiveInfinity, splits.Length + 1).ToArray();
            foreach (TextLayoutEngine.TextLine line in lines) {
                consumeWork?.Invoke(1);
                cancellationCheck?.Invoke();
                if (line.Y > upper || line.Y <= lower || Math.Abs(line.Y - y) <= 1.5D) continue;
                foreach (PdfTextSpan span in line.Spans) {
                    consumeWork?.Invoke(1);
                    cancellationCheck?.Invoke();
                    if (span.IsBold || string.IsNullOrWhiteSpace(span.Text) || span.X < left || span.X > right) continue;
                    // Before the first body row, a header/body gap independently
                    // bounds leading text even if that cell has no trailing line.
                    if (index == firstBodyIndex && line.Y > y) centered.Add(index);
                    double gap = Math.Abs(line.Y - y);
                    if (gap > fontSize * 1.8D) continue;
                    int column = 0;
                    while (column < splits.Length && span.X >= splits[column]) column++;
                    if (line.Y > y) above[column] = Math.Min(above[column], gap);
                    else below[column] = Math.Min(below[column], gap);
                }
            }
            for (int column = 0; column < above.Length; column++) {
                consumeWork?.Invoke(1);
                if (double.IsPositiveInfinity(above[column]) || double.IsPositiveInfinity(below[column])) continue;
                if (Math.Abs(above[column] - below[column]) <= Math.Max(1.5D, fontSize * 0.2D)) {
                    centered.Add(index);
                    break;
                }
            }
        }
        return centered;
    }
}
