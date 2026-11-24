using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Notebook.Markdown;

internal static partial class NotebookPlotlyChart {
    private static OfficeDrawing? RenderScene(SceneLayout layout, List<SceneTrace> traces, CancellationToken cancellationToken) {
        SceneDomain domain = SceneDomain.Create(layout, traces);
        SceneScale? colorbar = traces[0].Scale is { Show: true } scale ? scale : null;
        int legendCount = layout.ShowLegend ? traces.Count(trace => trace.Legend) : 0;
        bool bottomLegend = legendCount != 0 && (layout.Legend == OfficeChartLegendPosition.Bottom || layout.Width < 360);
        double rightReserve = colorbar != null ? 48 : legendCount != 0 && !bottomLegend ? 82 : 0;
        double bottomReserve = bottomLegend ? ((legendCount + 3) / 4) * 15 + 8 : 0;
        if (!SceneView.TryCreate(layout, domain, rightReserve, bottomReserve, out SceneView view, out SceneProjection[] corners))
            return null;

        var drawing = new OfficeDrawing(layout.Width, layout.Height);
        OfficeShape background = OfficeShape.Rectangle(layout.Width, layout.Height);
        background.FillColor = layout.Style.BackgroundColor;
        drawing.AddShape(background, 0, 0);
        if (!string.IsNullOrWhiteSpace(layout.Title)) {
            drawing.AddText(layout.Title!, 12, 10, layout.Width - 24, 29, new OfficeFontInfo(null, 16, OfficeFontStyle.Bold),
                layout.TextColor, OfficeTextAlignment.Center, shrinkToFit: true);
        }

        PaintSceneAxes(drawing, layout, view, corners);
        if (traces[0].Faces != null) {
            if (!PaintSceneSurface(drawing, traces[0], view, cancellationToken)) return null;
        } else if (!PaintSceneScatter(drawing, traces, view, cancellationToken)) {
            return null;
        }
        if (colorbar != null) PaintSceneColorbar(drawing, colorbar, layout.TextColor);
        if (legendCount != 0) PaintSceneLegend(drawing, traces, layout.TextColor, legendCount, bottomLegend, bottomReserve);
        return drawing;
    }

    private static bool PaintSceneSurface(OfficeDrawing drawing, SceneTrace trace, SceneView view, CancellationToken cancellationToken) {
        var points = new SceneProjection[trace.Points.Length];
        for (int index = 0; index < points.Length; index++) {
            if (!view.TryProject(trace.Points[index], out points[index])) return false;
        }
        var faces = new List<(double Depth, int Index)>(trace.Faces!.Count);
        for (int i = 0; i < trace.Faces.Count; i++) {
            SceneFace face = trace.Faces[i];
            faces.Add(((points[face.A].Depth + points[face.B].Depth + points[face.C].Depth) / 3, i));
        }
        // Paint distant faces first to approximate self-occlusion while retaining vector output.
        faces.Sort(static (left, right) => CompareSceneDepth(left.Depth, right.Depth, left.Index, right.Index));
        foreach (var item in faces) {
            cancellationToken.ThrowIfCancellationRequested();
            SceneFace face = trace.Faces[item.Index];
            PaintScenePolygon(drawing, trace.Scale!.At(face.ColorValue),
                view.Place(points[face.A]), view.Place(points[face.B]), view.Place(points[face.C]));
        }
        return true;
    }

    private static bool PaintSceneScatter(OfficeDrawing drawing, List<SceneTrace> traces, SceneView view, CancellationToken cancellationToken) {
        var marks = new List<SceneMark>();
        foreach (SceneTrace trace in traces) {
            SceneProjection previous = default;
            for (int index = 0; index < trace.Points.Length; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!view.TryProject(trace.Points[index], out SceneProjection point)) return false;
                if (trace.Lines && index > 0) {
                    marks.Add(new SceneMark((previous.Depth + point.Depth) / 2, marks.Count, trace, view.Place(previous), view.Place(point)));
                }
                if (trace.Markers) marks.Add(new SceneMark(point.Depth, marks.Count, trace, view.Place(point), null));
                previous = point;
            }
        }
        marks.Sort(static (left, right) => CompareSceneDepth(left.Depth, right.Depth, left.Index, right.Index));
        foreach (SceneMark mark in marks) {
            cancellationToken.ThrowIfCancellationRequested();
            PlotlyTraceStyle style = mark.Trace.Style;
            if (mark.End is OfficePoint end) {
                PaintSceneLine(drawing, mark.Start, end, style.LineColor ?? style.MarkerColor ?? mark.Trace.SeriesColor,
                    style.LineWidth?.ToPoints() ?? 1.7, style.LineDash ?? OfficeStrokeDashStyle.Solid);
            } else {
                PaintSceneMarker(drawing, mark.Start, style.MarkerColor ?? style.LineColor ?? mark.Trace.SeriesColor,
                    style.MarkerSize?.ToPoints() ?? 5.5, style.MarkerShape ?? OfficeChartMarkerShape.Circle);
            }
        }
        return true;
    }

    private static void PaintSceneMarker(OfficeDrawing drawing, OfficePoint point, OfficeColor color, double diameter, OfficeChartMarkerShape marker) {
        double radius = diameter / 2;
        double left = point.X - radius;
        double top = point.Y - radius;
        double right = point.X + radius;
        double bottom = point.Y + radius;
        if (marker == OfficeChartMarkerShape.Diamond) {
            PaintScenePolygon(drawing, color, new OfficePoint(left, point.Y), new OfficePoint(point.X, top),
                new OfficePoint(right, point.Y), new OfficePoint(point.X, bottom));
        } else if (marker == OfficeChartMarkerShape.X) {
            PaintSceneLine(drawing, new OfficePoint(left, top), new OfficePoint(right, bottom), color, Math.Max(0.8, radius / 2));
            PaintSceneLine(drawing, new OfficePoint(left, bottom), new OfficePoint(right, top), color, Math.Max(0.8, radius / 2));
        } else {
            OfficeShape shape = marker == OfficeChartMarkerShape.Square
                ? OfficeShape.Rectangle(diameter, diameter) : OfficeShape.Ellipse(diameter, diameter);
            shape.FillColor = color;
            drawing.AddShape(shape, left, top);
        }
    }

    private static int CompareSceneDepth(double left, double right, int leftIndex, int rightIndex) {
        int depth = right.CompareTo(left);
        return depth != 0 ? depth : leftIndex.CompareTo(rightIndex);
    }

    private static void PaintScenePolygon(OfficeDrawing drawing, OfficeColor color, params OfficePoint[] points) {
        double area = 0;
        for (int i = 0; i < points.Length; i++) {
            OfficePoint a = points[i];
            OfficePoint b = points[(i + 1) % points.Length];
            area += a.X * b.Y - b.X * a.Y;
        }
        if (Math.Abs(area) < 1e-7) return;
        OfficeShape polygon = OfficeShape.Polygon(points);
        polygon.FillColor = color;
        // A matching thin stroke covers antialias seams between adjacent faces.
        polygon.StrokeColor = color;
        polygon.StrokeWidth = 0.18;
        drawing.AddShape(polygon, points.Min(point => point.X), points.Min(point => point.Y));
    }

    private static void PaintSceneLine(OfficeDrawing drawing, OfficePoint start, OfficePoint end, OfficeColor color, double width,
        OfficeStrokeDashStyle dash = OfficeStrokeDashStyle.Solid) {
        if (Math.Abs(start.X - end.X) + Math.Abs(start.Y - end.Y) < 1e-7) return;
        OfficeShape line = OfficeShape.Line(start, end);
        line.StrokeColor = color;
        line.StrokeWidth = width;
        line.StrokeDashStyle = dash;
        drawing.AddShape(line, Math.Min(start.X, end.X), Math.Min(start.Y, end.Y));
    }

    private static OfficePoint MixScenePoint(OfficePoint a, OfficePoint b, double t) =>
        new(a.X + (b.X - a.X) * t, a.Y + (b.Y - a.Y) * t);

    private static string FormatSceneNumber(double value) => value.ToString("G4", CultureInfo.InvariantCulture);

    private readonly record struct SceneMark(double Depth, int Index, SceneTrace Trace, OfficePoint Start, OfficePoint? End);
}
