using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    /// <summary>
    /// Keeps real series vertices connected so explicit joins survive SVG, PDF and raster export.
    /// Clips original edges, never the resulting polygon boundary: a clip edge is not series ink.
    /// </summary>
    private static void AddJoinedSeriesOutline(OfficeDrawing drawing, IReadOnlyList<OfficePoint> points,
        OfficeColor color, double width, OfficeStrokeDashStyle dashStyle, OfficeStrokeLineJoin join,
        IReadOnlyList<bool>? plotted = null, bool closed = false,
        ChartPlotBounds? plotBounds = null, ValueRange? valueRange = null, ValueRange? xRange = null) {
        if (points.Count < 2) return;
        var runs = new List<List<OfficePoint>>();
        var run = new List<OfficePoint>();
        bool clipped = plotBounds.HasValue && valueRange.HasValue;

        bool Valid(int index) => (plotted == null || plotted[index]) &&
            IsFiniteChartValue(points[index].X) && IsFiniteChartValue(points[index].Y);

        bool Inside(OfficePoint point) => !clipped ||
            (point.Y >= valueRange!.Value.Min && point.Y <= valueRange.Value.Max &&
             (!xRange.HasValue || (point.X >= xRange.Value.Min && point.X <= xRange.Value.Max)));

        void Flush() {
            if (run.Count > 1) runs.Add(run);
            run = new List<OfficePoint>();
        }

        int edgeCount = closed ? points.Count : points.Count - 1;
        for (int edge = 0; edge < edgeCount; edge++) {
            int next = (edge + 1) % points.Count;
            if (!Valid(edge) || !Valid(next)) {
                Flush();
                continue;
            }
            OfficePoint start = points[edge], end = points[next];
            if (clipped) {
                bool visible = xRange.HasValue
                    ? TryClipScatterSegment(start, end, xRange.Value, valueRange!.Value, plotBounds!.Value, out start, out end)
                    : TryClipValueSegment(start, end, valueRange!.Value, plotBounds!.Value, out start, out end);
                if (!visible) {
                    Flush();
                    continue;
                }
                // Evaluate a retained vertex identically from either adjacent edge. Clipping
                // interpolation can otherwise differ by an ulp and accidentally split its join.
                if (Inside(points[edge])) start = MapVertex(points[edge]);
                if (Inside(points[next])) end = MapVertex(points[next]);
            }
            if (!IsFiniteChartValue(start.X) || !IsFiniteChartValue(start.Y) ||
                !IsFiniteChartValue(end.X) || !IsFiniteChartValue(end.Y)) {
                Flush();
                continue;
            }
            if (start != end) {
                if (run.Count > 0 && run[run.Count - 1] != start) Flush();
                if (run.Count == 0) run.Add(start);
                run.Add(end);
            }
            // Even coincident clipped endpoints must not join across an excursion outside the plot.
            if (!Inside(points[next])) Flush();
        }
        Flush();

        // The array seam is a real vertex of a closed contour, including partially clipped areas.
        if (closed && runs.Count > 1 && Valid(0) && Inside(points[0])) {
            List<OfficePoint> first = runs[0], last = runs[runs.Count - 1];
            if (last[last.Count - 1] == first[0]) {
                last.AddRange(first.GetRange(1, first.Count - 1));
                runs.RemoveAt(0);
            }
        }
        foreach (List<OfficePoint> path in runs) {
            bool closePath = closed && path.Count > 2 && path[0] == path[path.Count - 1];
            if (closePath) path.RemoveAt(path.Count - 1);
            double left = path[0].X, top = path[0].Y, right = left, bottom = top;
            foreach (OfficePoint point in path) {
                left = Math.Min(left, point.X);
                top = Math.Min(top, point.Y);
                right = Math.Max(right, point.X);
                bottom = Math.Max(bottom, point.Y);
            }
            var commands = new List<OfficePathCommand>(path.Count + 1);
            for (int index = 0; index < path.Count; index++) {
                var local = new OfficePoint(path[index].X - left, path[index].Y - top);
                commands.Add(index == 0 ? OfficePathCommand.MoveTo(local) : OfficePathCommand.LineTo(local));
            }
            if (closePath) commands.Add(OfficePathCommand.Close());
            // A horizontal/vertical open path still needs a positive declared canvas.
            OfficeShape shape = OfficeShape.Path(Math.Max(1D, right - left), Math.Max(1D, bottom - top), commands);
            shape.StrokeLineJoin = join;
            AddShape(drawing, shape, left, top, null, color, width, dashStyle);
        }

        OfficePoint MapVertex(OfficePoint point) => new OfficePoint(
            xRange.HasValue ? ToPlotX(point.X, xRange.Value.Min, xRange.Value.Max,
                plotBounds!.Value.Left, plotBounds.Value.Width) : point.X,
            ToPlotY(point.Y, valueRange!.Value.Min, valueRange.Value.Max,
                plotBounds!.Value.Top, plotBounds.Value.Height));
    }
}
