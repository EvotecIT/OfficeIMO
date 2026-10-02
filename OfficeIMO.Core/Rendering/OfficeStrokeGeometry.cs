using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Shared stroke-outline geometry for raster and native document painting.</summary>
internal static class OfficeStrokeGeometry {
    internal static IReadOnlyList<OfficeFlattenedPathContour> FlattenShape(OfficeShape shape, double pixelsPerUnit) {
        if (shape.Kind == OfficeShapeKind.Path) return OfficePathFlattener.Flatten(shape.PathCommands, 0D, 0D, 1D, pixelsPerUnit: pixelsPerUnit);
        IReadOnlyList<OfficePoint> points;
        bool closed = shape.Kind != OfficeShapeKind.Line;
        switch (shape.Kind) {
            case OfficeShapeKind.Line:
            case OfficeShapeKind.Polygon: points = shape.Points; break;
            case OfficeShapeKind.Ellipse:
                points = OfficeCurveFlattening.Ellipse(shape.Width/2,shape.Height/2,shape.Width/2,shape.Height/2,pixelsPerUnit); break;
            case OfficeShapeKind.RoundedRectangle:
                points = OfficeCurveFlattening.RoundedRectangle(0,0,shape.Width,shape.Height,shape.CornerRadius,shape.CornerRadius,pixelsPerUnit); break;
            default: points = new[] {new OfficePoint(0,0),new OfficePoint(shape.Width,0),new OfficePoint(shape.Width,shape.Height),new OfficePoint(0,shape.Height)}; break;
        }
        return new[] {new OfficeFlattenedPathContour(points,closed)};
    }

    internal static List<List<OfficePoint>> Create(IReadOnlyList<OfficeFlattenedPathContour> contours, double width,
        OfficeStrokeLineCap cap, OfficeStrokeLineJoin join, double miterLimit, IReadOnlyList<double>? dashPattern,
        double dashOffset, double pixelsPerUnit, double left, double top, double right, double bottom,
        CancellationToken cancellationToken = default, bool reset = false) {
        var outlines = new List<List<OfficePoint>>();
        if (!IsFinite(width) || width <= 0) return outlines;
        if (!IsFinite(miterLimit) || miterLimit < 1) miterLimit = 4;
        if (!IsFinite(pixelsPerUnit) || pixelsPerUnit <= 0) pixelsPerUnit = 1;
        IReadOnlyList<double> pattern = NormalizeRasterDashPattern(NormalizeDashPattern(dashPattern), .25 / pixelsPerUnit);
        double cycle = 0;
        foreach (double length in pattern) cycle = cycle > double.MaxValue - length ? double.MaxValue : cycle + length;
        foreach (OfficeFlattenedPathContour contour in contours) {
            cancellationToken.ThrowIfCancellationRequested();
            var points = new List<OfficePoint>(contour.Points);
            if (contour.Closed && points.Count > 1 && !SameStrokePoint(points[0], points[points.Count - 1])) points.Add(points[0]);
            foreach (List<OfficePoint> run in CollectStrokeRuns(points, width, miterLimit, pattern, cycle, dashOffset, reset, left, top, right, bottom, cancellationToken)) {
                cancellationToken.ThrowIfCancellationRequested();
                OfficeRasterStroker.AppendOutline(run, width, cap, join, miterLimit, pixelsPerUnit, outlines);
            }
        }
        return outlines;
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
    private static double Distance(double x1, double y1, double x2, double y2) => Math.Sqrt((x2-x1)*(x2-x1)+(y2-y1)*(y2-y1));

    private static bool TryClipStrokeLine(ref OfficePoint start, ref OfficePoint end, double width, double length,
        double left, double top, double right, double bottom, out double leading, out double trailing) {
        leading = trailing = 0;
        double dx = end.X-start.X, dy = end.Y-start.Y, first = 0, last = 1;
        double padding = Math.Max(1, width / 2 + 1);
        if (!Clip(-dx, start.X-left+padding, ref first, ref last) || !Clip(dx, right-start.X+padding, ref first, ref last)
            || !Clip(-dy, start.Y-top+padding, ref first, ref last) || !Clip(dy, bottom-start.Y+padding, ref first, ref last)) return false;
        OfficePoint original = start;
        start = new OfficePoint(original.X+dx*first,original.Y+dy*first);
        end = new OfficePoint(original.X+dx*last,original.Y+dy*last);
        leading = length * first; trailing = length * (1-last);
        return true;
    }

    private static bool Clip(double direction, double distance, ref double first, ref double last) {
        if (direction == 0) return distance >= 0;
        double fraction = distance / direction;
        if (direction < 0) { if (fraction > last) return false; first = Math.Max(first,fraction); }
        else { if (fraction < first) return false; last = Math.Min(last,fraction); }
        return true;
    }
    private static List<List<OfficePoint>> CollectStrokeRuns(IReadOnlyList<OfficePoint> points, double width,
        double miterLimit, IReadOnlyList<double> pattern, double cycle, double offset, bool reset, double left, double top, double right, double bottom, System.Threading.CancellationToken cancellationToken) {
        var runs = new List<List<OfficePoint>>();
        List<OfficePoint>? run = null;
        if (points.Count == 1) { runs.Add(new List<OfficePoint>(points)); return runs; }
        double phase = AdvancePatternPosition(0D, offset, cycle);
        int pieces = 0;
        for (int i = 1; i < points.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (reset) { phase = AdvancePatternPosition(0D, offset, cycle); run = null; }
            OfficePoint start = points[i - 1], end = points[i];
            double length = Distance(start.X, start.Y, end.X, end.Y);
            if (!IsFinite(length)) { run = null; continue; }
            if (length <= 1E-9D) {
                if (cycle <= 0D) AppendStrokeRun(runs, ref run, start, end);
                continue;
            }
            // Clip before subdividing dashes, preserving the phase of the invisible length.
            double paddingWidth = width * Math.Max(1D, miterLimit);
            if (!TryClipStrokeLine(ref start, ref end, paddingWidth, length, left, top, right, bottom, out double leading, out double trailing)) {
                phase = AdvancePatternPosition(phase, length, cycle); run = null; continue;
            }
            if (leading > 0D) run = null;
            phase = AdvancePatternPosition(phase, leading, cycle);
            double visibleLength = Distance(start.X, start.Y, end.X, end.Y);
            if (cycle <= 0D) {
                AppendStrokeRun(runs, ref run, start, end);
            } else {
                double position = 0D, patternOffset = phase;
                int index = 0;
                while (index < pattern.Count - 1 && patternOffset > 0D && patternOffset >= pattern[index]) { patternOffset -= pattern[index++]; }
                while (position < visibleLength) {
                    if (++pieces > 100000) throw new InvalidOperationException("The visible stroke exceeds the dash-piece limit.");
                    cancellationToken.ThrowIfCancellationRequested();
                    double remaining = pattern[index] - patternOffset;
                    if (remaining <= 0D) {
                        if (pattern[index] == 0D && (index & 1) == 0) {
                            var dot = new OfficePoint(start.X + (end.X - start.X) * (position / visibleLength), start.Y + (end.Y - start.Y) * (position / visibleLength));
                            runs.Add(new List<OfficePoint> { dot });
                        }
                        index = (index + 1) % pattern.Count; patternOffset = 0D;
                        continue;
                    }
                    double next = Math.Min(visibleLength, position + remaining);
                    if (next <= position) {
                        phase = AdvancePatternPosition(phase, remaining, cycle);
                        index = (index + 1) % pattern.Count; patternOffset = 0D;
                        continue;
                    }
                    if ((index & 1) == 0) {
                        var a = new OfficePoint(start.X + (end.X - start.X) * (position / visibleLength), start.Y + (end.Y - start.Y) * (position / visibleLength));
                        var b = new OfficePoint(start.X + (end.X - start.X) * (next / visibleLength), start.Y + (end.Y - start.Y) * (next / visibleLength));
                        AppendStrokeRun(runs, ref run, a, b);
                    } else run = null;
                    phase = AdvancePatternPosition(phase, next - position, cycle);
                    if (next - position >= remaining) { index = (index + 1) % pattern.Count; patternOffset = 0D; }
                    else patternOffset += next - position;
                    position = next;
                }
            }
            phase = AdvancePatternPosition(phase, trailing, cycle);
            if (trailing > 0D) run = null;
        }
        // A painted seam on a closed path is a join, not two caps.
        if (!reset && runs.Count > 1 && points.Count > 2 && SameStrokePoint(points[0], points[points.Count - 1])) {
            List<OfficePoint> last = runs[runs.Count - 1], first = runs[0];
            if (SameStrokePoint(last[last.Count - 1], first[0])) { last.AddRange(first.GetRange(1, first.Count - 1)); runs.RemoveAt(0); }
        }
        return runs;
    }

    private static void AppendStrokeRun(List<List<OfficePoint>> runs, ref List<OfficePoint>? run, OfficePoint start, OfficePoint end) {
        if (run == null || !SameStrokePoint(run[run.Count - 1], start)) { run = new List<OfficePoint> { start }; runs.Add(run); }
        run.Add(end);
    }

    private static bool SameStrokePoint(OfficePoint a, OfficePoint b) => Math.Abs(a.X - b.X) <= 1E-7D && Math.Abs(a.Y - b.Y) <= 1E-7D;
    private static List<double> NormalizeDashPattern(IReadOnlyList<double>? dashPattern) {
        List<double> pattern = new();
        if (dashPattern == null) {
            return pattern;
        }

        for (int i = 0; i < dashPattern.Count; i++) {
            double value = dashPattern[i];
            if (!IsFinite(value) || value < 0D) return new List<double>();
            pattern.Add(value);
        }

        if (!pattern.Exists(value => value > 0D)) return new List<double>();
        if ((pattern.Count & 1) == 1) {
            int originalCount = pattern.Count;
            for (int index = 0; index < originalCount; index++) pattern.Add(pattern[index]);
        }

        return pattern;
    }
    private static IReadOnlyList<double> NormalizeRasterDashPattern(IReadOnlyList<double> pattern, double minimum) {
        double smallest = double.MaxValue;
        for (int index = 0; index < pattern.Count; index++) {
            double length = pattern[index];
            if (IsFinite(length) && length > 0D) smallest = Math.Min(smallest, length);
        }
        if (smallest == double.MaxValue || smallest >= minimum) return pattern;

        double scale = minimum / smallest;
        var normalized = new double[pattern.Count];
        for (int index = 0; index < pattern.Count; index++) {
            normalized[index] = pattern[index] * scale;
            if (!IsFinite(scale) || !IsFinite(normalized[index])) {
                for (int fallbackIndex = 0; fallbackIndex < pattern.Count; fallbackIndex++) {
                    normalized[fallbackIndex] = Math.Max(pattern[fallbackIndex], minimum);
                }
                break;
            }
        }
        return normalized;
    }
    private static double AdvancePatternPosition(double patternPosition, double distance, double cycle) {
        if (!IsFinite(cycle) || cycle <= 0D) return 0D;
        double normalizedPosition = IsFinite(patternPosition) ? patternPosition % cycle : 0D;
        double normalizedDistance = IsFinite(distance) ? distance % cycle : 0D;
        double distanceUntilWrap = cycle - normalizedPosition;
        double advanced = normalizedDistance >= distanceUntilWrap
            ? normalizedDistance - distanceUntilWrap
            : normalizedPosition + normalizedDistance;
        return advanced < 0D ? advanced + cycle : advanced;
    }
}
