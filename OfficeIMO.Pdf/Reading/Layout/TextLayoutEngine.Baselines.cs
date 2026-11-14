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
        foreach (IReadOnlyList<PdfTextSpan> orientation in GroupByRotation(spans)) {
            cancellationCheck?.Invoke();
            var coordinates = new PdfTextBaselineCoordinates(orientation[0].RotationDegrees, useLogicalLayoutBaseline: true);
            List<PdfTextSpan> ordered = OrderByBaselineThenAlong(orientation, coordinates);
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
                    groups.Add(new BaselineSpanGroup(currentNormal, current));
                    current = new List<PdfTextSpan> { span };
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

    // Same groups, in the same order, as GroupBy(RotationKey). Pages are almost always one orientation,
    // which needs no grouping copy.
    private static IEnumerable<IReadOnlyList<PdfTextSpan>> GroupByRotation(IReadOnlyList<PdfTextSpan> spans) {
        if (spans.Count == 0) return Array.Empty<IReadOnlyList<PdfTextSpan>>();
        double first = RotationKey(spans[0]);
        for (int index = 1; index < spans.Count; index++) {
            if (!RotationKey(spans[index]).Equals(first)) {
                return spans.GroupBy(RotationKey).Select(static group => (IReadOnlyList<PdfTextSpan>)group.ToList());
            }
        }
        return new[] { spans };
    }

    // Stable order by descending baseline normal, then ascending position along it, as
    // OrderByDescending(Normal).ThenBy(Along) would produce, with each key computed once.
    private static List<PdfTextSpan> OrderByBaselineThenAlong(IReadOnlyList<PdfTextSpan> spans, PdfTextBaselineCoordinates coordinates) {
        int count = spans.Count;
        var normals = new double[count];
        var alongs = new double[count];
        var order = new int[count];
        for (int index = 0; index < count; index++) {
            normals[index] = coordinates.Normal(spans[index]);
            alongs[index] = coordinates.Along(spans[index]);
            order[index] = index;
        }
        Array.Sort(order, (left, right) => {
            int comparison = normals[right].CompareTo(normals[left]);
            if (comparison != 0) return comparison;
            comparison = alongs[left].CompareTo(alongs[right]);
            return comparison != 0 ? comparison : left.CompareTo(right);
        });
        var ordered = new List<PdfTextSpan>(count);
        for (int index = 0; index < count; index++) ordered.Add(spans[order[index]]);
        return ordered;
    }
}
