namespace OfficeIMO.Pdf;

internal static partial class TextLayoutEngine {
    // The parser computes directions from transformed unit vectors, so independently
    // positioned fragments of one baseline can differ by floating-point roundoff.
    private static double RotationKey(PdfTextSpan span) =>
        Math.Round((span.RotationDegrees % 360D + 360D) % 360D, 1) % 360D;

    private sealed class BaselineSpanGroup {
        internal readonly double Normal;
        internal readonly List<PdfTextSpan> Spans;
        internal List<PdfTextSpan>? OriginalSpans;

        internal BaselineSpanGroup(double normal, List<PdfTextSpan> spans) {
            Normal = normal;
            Spans = spans;
        }
    }

    private static List<TextLine> BuildBaselineLines(IReadOnlyList<PdfTextSpan> spans, Options options,
        Action<long>? consumeWork, Action? cancellationCheck) {
        var lines = new List<TextLine>();
        foreach (IGrouping<double, PdfTextSpan> orientation in spans.GroupBy(RotationKey)) {
            cancellationCheck?.Invoke();
            var coordinates = new PdfTextBaselineCoordinates(orientation.First().RotationDegrees, useLogicalLayoutBaseline: true);
            var ordered = orientation.OrderByDescending(coordinates.Normal).ThenBy(coordinates.Along).ToList();
            var groups = new List<BaselineSpanGroup>();
            var current = new List<PdfTextSpan>();
            double currentNormal = coordinates.Normal(ordered[0]);
            double currentFont = ordered[0].FontSize;
            foreach (PdfTextSpan span in ordered) {
                cancellationCheck?.Invoke();
                double normal = coordinates.Normal(span);
                if (current.Count == 0) { current.Add(span); currentNormal = normal; continue; }
                double tolerance = Math.Max(0.5D, Math.Min(options.LineMergeMaxPoints,
                    Math.Min(currentFont, span.FontSize) * options.LineMergeToleranceEm));
                if (Math.Abs(normal - currentNormal) <= tolerance) {
                    current.Add(span);
                    currentFont = (currentFont * (current.Count - 1) + span.FontSize) / current.Count;
                } else {
                    groups.Add(new BaselineSpanGroup(currentNormal, new List<PdfTextSpan>(current)));
                    current.Clear();
                    current.Add(span);
                    currentNormal = normal;
                    currentFont = span.FontSize;
                }
            }
            if (current.Count > 0) groups.Add(new BaselineSpanGroup(currentNormal, current));
            AttachInlineScriptRuns(groups, coordinates, options, consumeWork, cancellationCheck);
            foreach (BaselineSpanGroup group in groups) {
                cancellationCheck?.Invoke();
                if (group.Spans.Count > 0) AddBuiltLines(lines, group.Spans, options, group.OriginalSpans);
            }
        }
        return lines;
    }
}
