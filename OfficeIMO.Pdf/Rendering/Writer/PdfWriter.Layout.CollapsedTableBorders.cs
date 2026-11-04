namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>A nil edge suppresses its neighbour even when this cell has no visible border paint.</summary>
    private static bool HasCellBorderPaintOrConflict(PdfCellBorder? border, CollapsedTableBorders? collapsedBorders) =>
        HasRenderableCellBorder(border) || collapsedBorders != null && border != null &&
        (border.TopPriority == PdfCellBorderPriority.Nil || border.BottomPriority == PdfCellBorderPriority.Nil);

    /// <summary>Resolves source cell rules only against neighbours painted in the same physical table frame.</summary>
    private sealed class CollapsedTableBorders {
        private sealed class Edge {
            public Edge(bool top, PdfCellBorderPriority priority, double left, double right, double y,
                Action<StringBuilder, double, double> paint) {
                Top = top; Priority = priority; Left = left; Right = right; Y = y; Paint = paint;
            }
            public bool Top { get; }
            public PdfCellBorderPriority Priority { get; }
            public double Left { get; }
            public double Right { get; }
            public double Y { get; }
            public Action<StringBuilder, double, double> Paint { get; }
            public List<(double Left, double Right)> Hidden { get; } = new();
        }

        private readonly List<Edge> edges = new();

        public void Add(bool top, PdfCellBorderPriority priority, bool enabled, PdfCellBorderSide? side,
            double left, double right, double y, Action<StringBuilder, double, double> paint,
            HashSet<int>? hidden = null, double[]? widths = null) {
            if (right <= left || priority != PdfCellBorderPriority.Nil && (!enabled || !IsRenderableCellBorderSide(side))) return;
            if (hidden == null || hidden.Count == 0 || widths == null) {
                edges.Add(new Edge(top, priority, left, right, y, paint));
                return;
            }
            double position = left;
            for (int index = 0; index < widths.Length && position < right; index++) {
                double end = index == widths.Length - 1 ? right : Math.Min(right, position + widths[index]);
                if (!hidden.Contains(index) && end > position) edges.Add(new Edge(top, priority, position, end, y, paint));
                position = end;
            }
        }

        public void Flush(StringBuilder output) {
            if (edges.Count == 0) return;
            var ordered = edges.OrderBy(edge => edge.Y).ToArray();
            for (int start = 0; start < ordered.Length;) {
                int end = start + 1;
                while (end < ordered.Length && ordered[end].Y - ordered[start].Y < .001D) end++;
                var top = ordered.Skip(start).Take(end - start).Where(edge => edge.Top).OrderBy(edge => edge.Left).ToArray();
                var bottom = ordered.Skip(start).Take(end - start).Where(edge => !edge.Top).OrderBy(edge => edge.Left).ToArray();
                int i = 0, j = 0;
                while (i < top.Length && j < bottom.Length) {
                    Edge first = top[i], second = bottom[j];
                    double left = Math.Max(first.Left, second.Left), right = Math.Min(first.Right, second.Right);
                    if (right > left) {
                        if (Suppresses(first, second)) second.Hidden.Add((left, right));
                        if (Suppresses(second, first)) first.Hidden.Add((left, right));
                    }
                    if (first.Right <= second.Right) i++;
                    if (second.Right <= first.Right) j++;
                }
                start = end;
            }
            foreach (Edge edge in edges) {
                if (edge.Priority == PdfCellBorderPriority.Nil) continue;
                double left = edge.Left;
                foreach (var hidden in edge.Hidden.OrderBy(range => range.Left)) {
                    if (hidden.Left > left) edge.Paint(output, left, hidden.Left);
                    left = Math.Max(left, hidden.Right);
                }
                if (left < edge.Right) edge.Paint(output, left, edge.Right);
            }
            edges.Clear();
        }

        private static bool Suppresses(Edge winner, Edge other) =>
            winner.Priority == PdfCellBorderPriority.Nil ||
            winner.Priority == PdfCellBorderPriority.Cell && other.Priority == PdfCellBorderPriority.Table;
    }
}
