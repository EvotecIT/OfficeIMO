namespace OfficeIMO.Pdf;

internal static partial class TextLayoutEngine {
    // The parser computes directions from transformed unit vectors, so independently
    // positioned fragments of one baseline can differ by floating-point roundoff.
    private static double RotationKey(PdfTextSpan span) =>
        Math.Round((span.RotationDegrees % 360D + 360D) % 360D, 1) % 360D;

    private static List<TextLine> BuildBaselineLines(IReadOnlyList<PdfTextSpan> spans, Options options,
        Action? cancellationCheck) {
        var lines = new List<TextLine>();
        foreach (IGrouping<double, PdfTextSpan> orientation in spans.GroupBy(RotationKey)) {
            cancellationCheck?.Invoke();
            var coordinates = new PdfTextBaselineCoordinates(orientation.First().RotationDegrees);
            var ordered = orientation.OrderByDescending(coordinates.Normal).ThenBy(coordinates.Along).ToList();
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
                    AddBuiltLines(lines, current, options);
                    current.Clear();
                    current.Add(span);
                    currentNormal = normal;
                    currentFont = span.FontSize;
                }
            }
            if (current.Count > 0) AddBuiltLines(lines, current, options);
        }
        return lines;
    }
}
