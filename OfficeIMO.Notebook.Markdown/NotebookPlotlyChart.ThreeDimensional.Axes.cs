using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Notebook.Markdown;

internal static partial class NotebookPlotlyChart {
    private static void PaintSceneAxes(OfficeDrawing drawing, SceneLayout layout, SceneView view, SceneProjection[] projections) {
        var corners = new OfficePoint[projections.Length];
        int far = 0;
        for (int index = 0; index < corners.Length; index++) {
            corners[index] = view.Place(projections[index]);
            if (projections[index].Depth > projections[far].Depth) far = index;
        }
        OfficeColor wall = OfficeColor.FromRgb(236, 240, 248);
        OfficeColor edge = OfficeColor.FromRgb(169, 184, 204);
        OfficeColor grid = OfficeColor.FromRgb(211, 219, 231);
        for (int axis = 0; axis < 3; axis++) {
            int next = 1 << ((axis + 1) % 3);
            int other = 1 << ((axis + 2) % 3);
            PaintScenePolygon(drawing, wall, corners[far], corners[far ^ next], corners[far ^ (next | other)], corners[far ^ other]);
        }
        for (int axis = 0; axis < 3; axis++) {
            int next = 1 << ((axis + 1) % 3);
            int other = 1 << ((axis + 2) % 3);
            for (int tick = 1; tick < 4; tick++) {
                double t = tick / 4.0;
                OfficePoint first = ProjectSceneEdge(view, projections[far], projections[far ^ next], t);
                OfficePoint second = ProjectSceneEdge(view, projections[far ^ other], projections[far ^ (next | other)], t);
                PaintSceneLine(drawing, first, second, grid, 0.5);
                first = ProjectSceneEdge(view, projections[far], projections[far ^ other], t);
                second = ProjectSceneEdge(view, projections[far ^ next], projections[far ^ (next | other)], t);
                PaintSceneLine(drawing, first, second, grid, 0.5);
            }
        }

        string?[] labels = { layout.XTitle, layout.YTitle, layout.ZTitle };
        for (int axis = 0; axis < 3; axis++) {
            int bit = 1 << axis;
            int edgeStart = 0;
            double best = double.NegativeInfinity;
            for (int index = 0; index < corners.Length; index++) {
                if ((index & bit) != 0) continue;
                PaintSceneLine(drawing, corners[index], corners[index | bit], edge, 0.65);
                OfficePoint mid = MixScenePoint(corners[index], corners[index | bit], 0.5);
                double score = axis == 2 ? -mid.X : mid.Y;
                if (score > best) {
                    best = score;
                    edgeStart = index;
                }
            }
            PaintSceneAxisLabels(drawing, view, axis, projections[edgeStart], projections[edgeStart | bit], labels[axis], layout.TextColor);
        }
    }

    private static void PaintSceneAxisLabels(OfficeDrawing drawing, SceneView view, int axis, SceneProjection a, SceneProjection b,
        string? name, OfficeColor textColor) {
        OfficePoint first = view.Place(a);
        OfficePoint last = view.Place(b);
        double dx = last.X - first.X;
        double dy = last.Y - first.Y;
        double length = Math.Sqrt(dx * dx + dy * dy);
        if (length < 20) return;
        OfficePoint middle = ProjectSceneEdge(view, a, b, 0.5);
        double normalX = -dy / length;
        double normalY = dx / length;
        if (normalX * (middle.X - view.CenterX) + normalY * (middle.Y - view.CenterY) < 0) {
            normalX = -normalX;
            normalY = -normalY;
        }
        for (int index = 0; index < 3; index++) {
            double t = index / 2.0;
            OfficePoint tick = ProjectSceneEdge(view, a, b, t);
            PaintSceneLabel(drawing, FormatSceneNumber(view.Domain.AxisValue(axis, t)),
                tick.X + normalX * 14, tick.Y + normalY * 10, 42, 9, textColor);
        }
        if (string.IsNullOrWhiteSpace(name)) return;
        if (Math.Abs(dy) > Math.Abs(dx)) {
            // Keep the horizontal title on its own line beside a mostly vertical axis.
            OfficePoint title = ProjectSceneEdge(view, a, b, dy > 0 ? 0.25 : 0.75);
            PaintSceneLabel(drawing, name!, title.X + normalX * 33, title.Y, 68, 10, textColor);
        } else {
            PaintSceneLabel(drawing, name!, middle.X + normalX * 25, middle.Y + normalY * 23, 90, 10, textColor);
        }
    }

    private static OfficePoint ProjectSceneEdge(SceneView view, SceneProjection a, SceneProjection b, double t) {
        // Project a data-space fraction along a 3D edge, preserving perspective foreshortening.
        double fraction = view.Camera.Orthographic ? t : t * b.Depth / ((1 - t) * a.Depth + t * b.Depth);
        return MixScenePoint(view.Place(a), view.Place(b), fraction);
    }

    private static void PaintSceneColorbar(OfficeDrawing drawing, SceneScale scale, OfficeColor textColor) {
        const int steps = 24;
        double left = drawing.Width - 51;
        double top = drawing.Height * 0.28;
        double height = drawing.Height * 0.43;
        double step = height / steps;
        for (int index = 0; index < steps; index++) {
            var fill = OfficeShape.Rectangle(9, step + 0.1);
            fill.FillColor = scale.FromFraction(1 - (index + 0.5) / steps);
            drawing.AddShape(fill, left, top + index * step);
        }
        PaintSceneLabel(drawing, FormatSceneNumber(scale.Max), left + 25, top, 31, 8, textColor);
        PaintSceneLabel(drawing, FormatSceneNumber((scale.Min + scale.Max) / 2), left + 25, top + height / 2, 31, 8, textColor);
        PaintSceneLabel(drawing, FormatSceneNumber(scale.Min), left + 25, top + height, 31, 8, textColor);
    }

    private static void PaintSceneLegend(OfficeDrawing drawing, List<SceneTrace> traces, OfficeColor textColor, int count,
        bool bottomPosition, double bottomReserve) {
        int shown = 0;
        double maxWidth = bottomPosition ? (drawing.Width - 24) / Math.Min(4, count) : 74;
        for (int index = 0; index < traces.Count; index++) {
            SceneTrace trace = traces[index];
            if (!trace.Legend) continue;
            double x = bottomPosition ? 12 + shown % 4 * maxWidth : drawing.Width - 81;
            double y = bottomPosition ? drawing.Height - bottomReserve + 7 + shown / 4 * 15 : 52 + shown * 15;
            OfficeColor paint = trace.Style.MarkerColor ?? trace.Style.LineColor ?? trace.SeriesColor;
            if (trace.Lines) PaintSceneLine(drawing, new OfficePoint(x, y), new OfficePoint(x + 9, y), paint, 1.5);
            if (trace.Markers) PaintSceneMarker(drawing, new OfficePoint(x + 4.5, y), paint, 4.5,
                trace.Style.MarkerShape ?? OfficeChartMarkerShape.Circle);
            string label = trace.Name ?? "trace " + index.ToString(CultureInfo.InvariantCulture);
            drawing.AddText(label, x + 13, y - 6, maxWidth - 13, 12, new OfficeFontInfo(null, 9), textColor, shrinkToFit: true);
            shown++;
        }
    }

    private static void PaintSceneLabel(OfficeDrawing drawing, string label, double x, double y, double width,
        double fontSize, OfficeColor color) {
        double height = fontSize + 3;
        double left = Math.Max(0, Math.Min(drawing.Width - width, x - width / 2));
        double top = Math.Max(0, Math.Min(drawing.Height - height, y - height / 2));
        drawing.AddText(label, left, top, width, height, new OfficeFontInfo(null, fontSize), color,
            OfficeTextAlignment.Center, shrinkToFit: true);
    }
}
