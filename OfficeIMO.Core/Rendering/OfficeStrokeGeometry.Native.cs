using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

internal static partial class OfficeStrokeGeometry {
    // Build native stroke coverage as one nonzero union. Endpoint caps supplement
    // dash caps only where the authored endpoint touches a painted dash.
    internal static List<List<OfficePoint>> CreateNative(IReadOnlyList<OfficeFlattenedPathContour> contours,
        double width, OfficeStrokeLineJoin join, double miterLimit, IReadOnlyList<double>? dashPattern,
        double dashOffset, double pixelsPerUnit, double left, double top, double right, double bottom,
        OfficeStrokeOutlineOptions options) {
        var outlines = new List<List<OfficePoint>>();
        if (!IsFinite(width) || width <= 0D) return outlines;
        options = CopyNativeStyle(options, options.StartCap, options.EndCap);
        IReadOnlyList<double> pattern = NormalizeDashPattern(dashPattern);
        double cycle = 0D;
        foreach (double length in pattern) cycle += length;
        if (!IsFinite(cycle)) throw new InvalidOperationException("Native stroke dash cycle is not finite.");
        foreach (var contour in contours) {
            options.CancellationToken.ThrowIfCancellationRequested();
            var points = new List<OfficePoint>(contour.Points);
            if (points.Count == 0 || points.Count == 1 && !contour.Closed) continue;
            if (points.Count == 1) points.Add(points[0]);
            if (contour.Closed && !points[0].Equals(points[points.Count - 1])) points.Add(points[0]);
            double length = 0D;
            for (int i = 1; i < points.Count; i++) length += Distance(points[i-1].X, points[i-1].Y, points[i].X, points[i].Y);
            if (!IsFinite(length)) throw new InvalidOperationException("Native stroke path length is not finite.");
            if (length == 0D) {
                var dotStyle = CopyNativeStyle(options, contour.UseStartLineCap ? options.StartCap : OfficeStrokeOutlineCap.Flat,
                    contour.UseEndLineCap ? options.EndCap : OfficeStrokeOutlineCap.Flat);
                if (contour.Closed) dotStyle.StartCap = dotStyle.EndCap = OfficeStrokeOutlineCap.Round;
                OfficeRasterStroker.AppendOutline(points, width, join, miterLimit, pixelsPerUnit, outlines, dotStyle);
                continue;
            }
            var dotDirections = new Dictionary<List<OfficePoint>, OfficePoint>();
            var closedRuns = new HashSet<List<OfficePoint>>();
            foreach (var run in CollectStrokeRuns(points, width, miterLimit, pattern, cycle, dashOffset, false,
                left, top, right, bottom, options.CancellationToken, dotDirections, true, contour.Closed, closedRuns)) {
                if (dotDirections.TryGetValue(run, out var direction)) {
                    OfficeRasterStroker.AddCap(outlines, run[0], direction, width / 2D, options.DashCap, pixelsPerUnit, options);
                    OfficeRasterStroker.AddCap(outlines, run[0], new OfficePoint(-direction.X, -direction.Y), width / 2D, options.DashCap, pixelsPerUnit, options);
                    continue;
                }
                var start = cycle > 0D ? options.DashCap :
                    contour.UseStartLineCap && run[0].Equals(points[0]) ? options.StartCap : OfficeStrokeOutlineCap.Flat;
                var end = cycle > 0D ? options.DashCap :
                    contour.UseEndLineCap && run[run.Count - 1].Equals(points[points.Count - 1]) ? options.EndCap : OfficeStrokeOutlineCap.Flat;
                var style = CopyNativeStyle(options, start, end);
                style.Closed = closedRuns.Contains(run);
                OfficeRasterStroker.AppendOutline(run, width, join, miterLimit, pixelsPerUnit, outlines,
                    style);
            }
            if (cycle > 0D && !contour.Closed) {
                if (contour.UseStartLineCap && TouchesPaintedDash(dashOffset, pattern, cycle))
                    AddNativeEndpointCap(points, false, width, pixelsPerUnit, options.StartCap, options, outlines);
                if (contour.UseEndLineCap && TouchesPaintedDash(AdvancePatternPosition(dashOffset, length, cycle), pattern, cycle))
                    AddNativeEndpointCap(points, true, width, pixelsPerUnit, options.EndCap, options, outlines);
            }
        }
        return outlines;
    }

    private static OfficeStrokeOutlineOptions CopyNativeStyle(OfficeStrokeOutlineOptions options,
        OfficeStrokeOutlineCap start, OfficeStrokeOutlineCap end) => new() {
            StartCap = start, EndCap = end, DashCap = options.DashCap, ClipMiter = options.ClipMiter,
            DegenerateMiterLimitOne = options.DegenerateMiterLimitOne, CancellationToken = options.CancellationToken,
            Closed = false, PreserveExactGeometry = true,
            ChargePoints = options.ChargePoints
        };

    private static bool TouchesPaintedDash(double phase, IReadOnlyList<double> pattern, double cycle) {
        phase = AdvancePatternPosition(0D, phase, cycle);
        double position = 0D;
        for (int i = 0; i < pattern.Count; i++) {
            double end = position + pattern[i];
            double tolerance = cycle * 1E-12D;
            if ((i & 1) == 0 && phase >= position - tolerance && phase <= end + tolerance) return true;
            position = end;
        }
        return false;
    }

    private static void AddNativeEndpointCap(IReadOnlyList<OfficePoint> points, bool atEnd, double width,
        double pixelsPerUnit, OfficeStrokeOutlineCap cap, OfficeStrokeOutlineOptions options, List<List<OfficePoint>> output) {
        if (cap == OfficeStrokeOutlineCap.Flat) return;
        int endpoint = atEnd ? points.Count - 1 : 0;
        int step = atEnd ? -1 : 1;
        var end = points[endpoint];
        for (int i = endpoint + step; i >= 0 && i < points.Count; i += step) {
            double dx = end.X - points[i].X, dy = end.Y - points[i].Y;
            double length = Math.Sqrt(dx * dx + dy * dy);
            if (length == 0D) continue;
            OfficeRasterStroker.AddCap(output, end, new OfficePoint(dx / length, dy / length), width / 2D, cap, pixelsPerUnit, options);
            return;
        }
    }
}
