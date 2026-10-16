using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeConicGradient {
    /// <summary>
    /// Expands the gradient into clipped solid-color vector regions shared by raster, SVG, and PDF exporters.
    /// Authored stop boundaries are retained in addition to the uniform quality segments.
    /// </summary>
    public OfficeDrawing CreateDrawing(double width, double height, int qualitySegments = 360) {
        ValidatePositive(width, nameof(width));
        ValidatePositive(height, nameof(height));
        if (qualitySegments < 12 || qualitySegments > 4096) {
            throw new ArgumentOutOfRangeException(nameof(qualitySegments), "Conic gradient quality must be between 12 and 4096 segments.");
        }

        var content = new OfficeDrawing(width, height);
        OfficeColor firstColor = Stops[0].Color;
        bool uniformColor = true;
        bool uniformAlpha = true;
        foreach (OfficeGradientStop stop in Stops) {
            uniformColor &= stop.Color == firstColor;
            uniformAlpha &= stop.Color.A == firstColor.A;
        }
        if (uniformColor) {
            if (firstColor.A > 0) {
                OfficeShape fill = OfficeShape.Rectangle(width, height);
                fill.FillColor = firstColor;
                fill.StrokeWidth = 0D;
                content.AddShape(fill, 0D, 0D);
            }
            return ClipPaint(content, width, height, 255);
        }
        byte paintAlpha = uniformAlpha ? firstColor.A : (byte)255;
        if (paintAlpha == 0) return ClipPaint(content, width, height, 255);

        var boundaries = new SortedSet<double>();
        for (int index = 0; index <= qualitySegments; index++) boundaries.Add((double)index / qualitySegments);
        foreach (OfficeGradientStop stop in Stops) boundaries.Add(stop.Offset);
        var ordered = new List<double>(boundaries);
        double centerX = CenterX * width;
        double centerY = CenterY * height;
        double cornerRadius = Math.Max(
            Math.Max(Distance(centerX, centerY, 0D, 0D), Distance(centerX, centerY, width, 0D)),
            Math.Max(Distance(centerX, centerY, 0D, height), Distance(centerX, centerY, width, height)));
        double segmentAngle = (Math.PI * 2D) / qualitySegments;
        double overlap = Math.Min(
            segmentAngle * 0.49D,
            Math.Atan2(1.25D, Math.Max(1D, cornerRadius)));
        double radius = cornerRadius / Math.Cos((Math.PI / qualitySegments) + overlap) * 1.000001D;
        OfficeDrawing? opacityField = uniformAlpha ? null : new OfficeDrawing(width, height);
        AddCoveragePaint(content, opacityField, width, height, centerX, centerY, cornerRadius, qualitySegments);
        var colorRun = new ConicPaintRun(content);
        ConicPaintRun? opacityRun = opacityField == null ? null : new ConicPaintRun(opacityField);
        for (int index = 1; index < ordered.Count; index++) {
            double start = ordered[index - 1];
            double end = ordered[index];
            if (end <= start) continue;
            // Flat regions already have their final paint. Painting them again
            // would bias antialiased hard edges without improving the field.
            if (!IsInterpolatingInterval(end)) {
                colorRun.Complete();
                opacityRun?.Complete();
                continue;
            }
            OfficeColor color = Sample((start + end) / 2D);

            // Retain overlap between quality rays, but never cross an authored
            // hard stop or the discontinuity at the start of a turn.
            double startOverlap = IsHardColorBoundary(start) ? 0D : overlap;
            double endOverlap = IsHardColorBoundary(end) ? 0D : overlap;
            double startRadians = ToRadians(StartAngle + (start * 360D)) - (Math.PI / 2D) - startOverlap;
            double endRadians = ToRadians(StartAngle + (end * 360D)) - (Math.PI / 2D) + endOverlap;
            double firstX = centerX + (Math.Cos(startRadians) * radius);
            double firstY = centerY + (Math.Sin(startRadians) * radius);
            double secondX = centerX + (Math.Cos(endRadians) * radius);
            double secondY = centerY + (Math.Sin(endRadians) * radius);
            List<OfficePoint> clippedWedge = ClipTriangleToRectangle(
                new OfficePoint(centerX, centerY),
                new OfficePoint(firstX, firstY),
                new OfficePoint(secondX, secondY), width, height);
            if (!HasPolygonArea(clippedWedge)) continue;

            colorRun.Add(clippedWedge, OfficeColor.FromRgb(color.R, color.G, color.B));
            opacityRun?.Add(clippedWedge, OfficeColor.FromRgb(color.A, color.A, color.A));
        }
        colorRun.Complete();
        opacityRun?.Complete();
        return ClipPaint(content, width, height, paintAlpha, opacityField);
    }

    /// <summary>
    /// Supplies complete opaque color/opacity fields with bounded cumulative angular regions.
    /// Fine wedges refine interpolating regions; their antialiased internal edges never reveal the backdrop.
    /// </summary>
    private void AddCoveragePaint(OfficeDrawing colorField, OfficeDrawing? opacityField,
        double width, double height, double centerX, double centerY, double cornerRadius, int qualitySegments) {
        double phase = 0D;
        foreach (OfficeGradientStop stop in Stops) {
            if (stop.Offset > 0D && stop.Offset < 1D && IsHardColorBoundary(stop.Offset)) {
                phase = stop.Offset;
                break;
            }
        }
        var boundaries = new SortedSet<double>();
        int coarseSegments = Math.Min(qualitySegments, 32);
        for (int index = 0; index <= coarseSegments; index++) boundaries.Add((double)index / coarseSegments);
        foreach (OfficeGradientStop stop in Stops) {
            // Fine refinement skips flat spans, so their endpoints must also
            // delimit the underlying color and opacity fields.
            double relative = stop.Offset - phase;
            if (relative < 0D) relative += 1D;
            if (relative >= 1D) relative -= 1D;
            boundaries.Add(relative);
        }
        var ordered = new List<double>(boundaries);
        OfficeColor previous = SampleCoverage(phase, (ordered[0] + ordered[1]) / 2D);
        AddOpaqueBox(colorField, width, height, OfficeColor.FromRgb(previous.R, previous.G, previous.B));
        if (opacityField != null) AddOpaqueBox(opacityField, width, height, OfficeColor.FromRgb(previous.A, previous.A, previous.A));
        // A suffix is split at quarter turns. Its outer chord must lie beyond
        // every box corner even when the gradient center is outside the box.
        double radius = cornerRadius / Math.Cos(Math.PI / 4D) * 1.000001D;
        for (int index = 1; index < ordered.Count - 1; index++) {
            double start = ordered[index];
            OfficeColor color = SampleCoverage(phase, (start + ordered[index + 1]) / 2D);
            OfficeColor rgb = OfficeColor.FromRgb(color.R, color.G, color.B);
            OfficeColor previousRgb = OfficeColor.FromRgb(previous.R, previous.G, previous.B);
            bool addColor = rgb != previousRgb;
            bool addOpacity = opacityField != null && color.A != previous.A;
            if (addColor || addOpacity) {
                List<OfficePathCommand> commands = CreateSuffixCommands(start, phase, width, height, centerX, centerY, radius,
                    out double x, out double y);
                if (addColor) AddPaintRun(colorField, commands, rgb, x, y);
                if (addOpacity) AddPaintRun(opacityField!, commands, OfficeColor.FromRgb(color.A, color.A, color.A), x, y);
            }
            previous = color;
        }
    }

    private OfficeColor SampleCoverage(double phase, double relative) {
        double offset = phase + relative;
        return Sample(offset >= 1D ? offset - 1D : offset);
    }

    private static void AddOpaqueBox(OfficeDrawing content, double width, double height, OfficeColor color) {
        OfficeShape fill = OfficeShape.Rectangle(width, height);
        fill.FillColor = color;
        fill.StrokeWidth = 0D;
        content.AddShape(fill, 0D, 0D);
    }

    /// <summary>Builds one clockwise suffix as a nonzero union of at most four clipped quarter-turn fans.</summary>
    private List<OfficePathCommand> CreateSuffixCommands(double start, double phase, double width, double height,
        double centerX, double centerY, double radius, out double x, out double y) {
        var boundaries = new List<double> { start };
        for (int quarter = 1; quarter <= 4; quarter++) {
            double boundary = quarter / 4D;
            if (boundary > start) boundaries.Add(boundary);
        }
        var commands = new List<OfficePathCommand>();
        x = y = double.PositiveInfinity;
        for (int index = 1; index < boundaries.Count; index++) {
            double first = ToRadians(StartAngle + (phase + boundaries[index - 1]) * 360D) - Math.PI / 2D;
            double second = ToRadians(StartAngle + (phase + boundaries[index]) * 360D) - Math.PI / 2D;
            List<OfficePoint> points = ClipTriangleToRectangle(new OfficePoint(centerX, centerY),
                new OfficePoint(centerX + Math.Cos(first) * radius, centerY + Math.Sin(first) * radius),
                new OfficePoint(centerX + Math.Cos(second) * radius, centerY + Math.Sin(second) * radius), width, height);
            if (!HasPolygonArea(points)) continue;
            commands.Add(OfficePathCommand.MoveTo(points[0].X, points[0].Y));
            for (int pointIndex = 0; pointIndex < points.Count; pointIndex++) {
                OfficePoint point = points[pointIndex];
                x = Math.Min(x, point.X); y = Math.Min(y, point.Y);
                if (pointIndex > 0) commands.Add(OfficePathCommand.LineTo(point.X, point.Y));
            }
            commands.Add(OfficePathCommand.Close());
        }
        return commands;
    }

    private static bool HasPolygonArea(IReadOnlyList<OfficePoint> points) {
        if (points.Count < 3) return false;
        double area = 0D;
        OfficePoint previous = points[points.Count - 1];
        foreach (OfficePoint point in points) {
            area += previous.X * point.Y - point.X * previous.Y;
            previous = point;
        }
        return area != 0D;
    }

    // Uniform quality boundaries include every stop, so the interval ending
    // here lies entirely within one original interpolation span.
    private bool IsInterpolatingInterval(double end) {
        for (int index = 1; index < Stops.Count; index++) {
            if (end <= Stops[index].Offset) return Stops[index - 1].Color != Stops[index].Color;
        }
        return false;
    }

    /// <summary>Paints the union of identically colored contours once, including their fractional coverage.</summary>
    private static void AddPaintRun(OfficeDrawing content, List<OfficePathCommand> commands, OfficeColor color, double x, double y) {
        if (commands.Count == 0 || color.A == 0) return;
        OfficeShape region = OfficeShape.Path(commands);
        region.FillColor = color;
        region.FillRule = OfficeFillRule.NonZero;
        region.StrokeWidth = 0D;
        content.AddShape(region, x, y);
    }

    /// <summary>Retains the paint box and applies alpha after overlapping opaque vector paint is complete.</summary>
    private static OfficeDrawing ClipPaint(OfficeDrawing content, double width, double height, byte alpha, OfficeDrawing? opacityField = null) {
        if (opacityField != null) {
            var isolated = new OfficeDrawing(width, height);
            OfficeDrawingSoftMask mask = OfficeDrawingSoftMask.CreateWithLuminosityStandard(
                opacityField, OfficeSoftMaskLuminosityStandard.PdfDeviceRgb, OfficeSoftMaskMode.Luminosity);
            isolated.AddEffectDrawing(content, OfficeTransform.Identity, OfficeBlendMode.Normal, mask);
            content = isolated;
        } else if (alpha < 255) {
            var isolated = new OfficeDrawing(width, height);
            isolated.AddEffectDrawing(content, OfficeTransform.Identity, alpha / 255D);
            content = isolated;
        }
        var drawing = new OfficeDrawing(width, height);
        drawing.AddClippedDrawing(content, 0D, 0D, OfficeClipPath.Rectangle(width, height));
        return drawing;
    }

    /// <summary>Collects one consecutive color region without changing its order among other regions.</summary>
    private sealed class ConicPaintRun {
        private readonly OfficeDrawing _content;
        private readonly List<OfficePathCommand> _commands = new List<OfficePathCommand>();
        private OfficeColor? _color;
        private double _x = double.PositiveInfinity, _y = double.PositiveInfinity;

        internal ConicPaintRun(OfficeDrawing content) => _content = content;

        internal void Add(IReadOnlyList<OfficePoint> points, OfficeColor color) {
            if (_color.HasValue && _color.Value != color) {
                Complete();
            }
            _color = color;
            _commands.Add(OfficePathCommand.MoveTo(points[0].X, points[0].Y));
            for (int index = 0; index < points.Count; index++) {
                OfficePoint point = points[index];
                _x = Math.Min(_x, point.X);
                _y = Math.Min(_y, point.Y);
                if (index > 0) _commands.Add(OfficePathCommand.LineTo(point.X, point.Y));
            }
            _commands.Add(OfficePathCommand.Close());
        }

        internal void Complete() {
            if (_color.HasValue) AddPaintRun(_content, _commands, _color.Value, _x, _y);
            _commands.Clear();
            _color = null;
            _x = _y = double.PositiveInfinity;
        }
    }

    private bool IsHardColorBoundary(double offset) {
        if (offset == 0D || offset == 1D) return Stops[0].Color != Stops[Stops.Count - 1].Color;
        for (int index = 1; index < Stops.Count; index++) {
            if (Stops[index].Offset == offset && Stops[index - 1].Offset == offset &&
                Stops[index].Color != Stops[index - 1].Color) return true;
        }
        return false;
    }

    private static double ToRadians(double degrees) => degrees * Math.PI / 180D;

    private static List<OfficePoint> ClipTriangleToRectangle(
        OfficePoint first,
        OfficePoint second,
        OfficePoint third,
        double width,
        double height) {
        var points = new List<OfficePoint> { first, second, third };
        points = ClipAgainstVerticalBoundary(points, 0D, keepGreater: true);
        points = ClipAgainstVerticalBoundary(points, width, keepGreater: false);
        points = ClipAgainstHorizontalBoundary(points, 0D, keepGreater: true);
        return ClipAgainstHorizontalBoundary(points, height, keepGreater: false);
    }

    private static double Distance(double x1, double y1, double x2, double y2) {
        double x = x2 - x1;
        double y = y2 - y1;
        return Math.Sqrt((x * x) + (y * y));
    }

    private static List<OfficePoint> ClipAgainstVerticalBoundary(
        IReadOnlyList<OfficePoint> input,
        double boundary,
        bool keepGreater) {
        var output = new List<OfficePoint>(input.Count + 1);
        if (input.Count == 0) return output;
        OfficePoint previous = input[input.Count - 1];
        bool previousInside = keepGreater ? previous.X >= boundary : previous.X <= boundary;
        for (int index = 0; index < input.Count; index++) {
            OfficePoint current = input[index];
            bool currentInside = keepGreater ? current.X >= boundary : current.X <= boundary;
            if (currentInside != previousInside) {
                double ratio = (boundary - previous.X) / (current.X - previous.X);
                output.Add(new OfficePoint(boundary, previous.Y + ((current.Y - previous.Y) * ratio)));
            }
            if (currentInside) output.Add(current);
            previous = current;
            previousInside = currentInside;
        }
        return output;
    }

    private static List<OfficePoint> ClipAgainstHorizontalBoundary(
        IReadOnlyList<OfficePoint> input,
        double boundary,
        bool keepGreater) {
        var output = new List<OfficePoint>(input.Count + 1);
        if (input.Count == 0) return output;
        OfficePoint previous = input[input.Count - 1];
        bool previousInside = keepGreater ? previous.Y >= boundary : previous.Y <= boundary;
        for (int index = 0; index < input.Count; index++) {
            OfficePoint current = input[index];
            bool currentInside = keepGreater ? current.Y >= boundary : current.Y <= boundary;
            if (currentInside != previousInside) {
                double ratio = (boundary - previous.Y) / (current.Y - previous.Y);
                output.Add(new OfficePoint(previous.X + ((current.X - previous.X) * ratio), boundary));
            }
            if (currentInside) output.Add(current);
            previous = current;
            previousInside = currentInside;
        }
        return output;
    }

}
