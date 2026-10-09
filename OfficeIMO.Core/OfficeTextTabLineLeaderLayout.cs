using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

// Geometry is planned once alongside the measured tab advance. Renderers consume
// the same filled outlines, including their bounds, color and resource limit.
internal sealed class OfficeTextTabLineLeaderPaint {
    internal OfficeTextTabLineLeaderPaint(List<IReadOnlyList<OfficePoint>> contours, OfficeColor color) {
        Contours = contours; Color = color;
        Left = Top = double.PositiveInfinity; Right = Bottom = double.NegativeInfinity;
        foreach (var contour in contours) foreach (OfficePoint point in contour) {
            Left = Math.Min(Left, point.X); Right = Math.Max(Right, point.X);
            Top = Math.Min(Top, point.Y); Bottom = Math.Max(Bottom, point.Y);
        }
    }
    internal IReadOnlyList<IReadOnlyList<OfficePoint>> Contours { get; }
    internal OfficeColor Color { get; }
    internal double Left { get; }
    internal double Right { get; }
    internal double Top { get; }
    internal double Bottom { get; }
    internal OfficeShape ToShape() {
        var commands = new List<OfficePathCommand>();
        foreach (var contour in Contours) {
            commands.Add(OfficePathCommand.MoveTo(contour[0]));
            for (int i = 1; i < contour.Count; i++) commands.Add(OfficePathCommand.LineTo(contour[i]));
            commands.Add(OfficePathCommand.Close());
        }
        OfficeShape shape = OfficeShape.Path(commands);
        shape.FillColor = Color; shape.StrokeColor = null; shape.StrokeWidth = 0;
        return shape;
    }
}

internal static class OfficeTextTabLineLeaderLayout {
    internal const int MaximumFrameVertices = 100000;
    internal const int MaximumTabVertices = 8192;

    internal static OfficeTextTabLineLeaderPaint? Create(OfficeTextTabLineLeader? leader, double gap, double fontSize,
        OfficeColor activeColor, OfficeTextTabLeaderLayout.Budget budget, CancellationToken cancellationToken, out bool limited) {
        limited = false; cancellationToken.ThrowIfCancellationRequested();
        if (leader == null || leader.Style == OfficeTextTabLineLeaderStyle.None || gap <= 0) return null;
        OfficeColor color = leader.Color ?? activeColor;
        if (color.A == 0) return null;
        double width = leader.WidthPoints ?? fontSize * leader.WidthFontFraction;
        if (!FinitePositive(width) || !FinitePositive(gap) || !FinitePositive(fontSize)) { limited = true; return null; }
        int available = Math.Min(MaximumTabVertices, budget.RemainingVertices);
        int used = 0; bool exhausted = false;
        var contours = new List<IReadOnlyList<OfficePoint>>();
        double amplitude = leader.Style == OfficeTextTabLineLeaderStyle.Wave ? Math.Max(width, fontSize / 12) : 0;
        double center = fontSize * .1D;
        double offset = leader.DoubleLine ? width + amplitude : 0;
        if (!FinitePositive(width * 20) || !FinitePositive(amplitude + width)) { limited = true; return null; }
        PaintRow(center - offset);
        if (leader.DoubleLine && !exhausted) PaintRow(center + offset);
        budget.RemainingVertices -= used; limited = exhausted;
        return contours.Count == 0 ? null : new OfficeTextTabLineLeaderPaint(contours, color);

        bool Add(IReadOnlyList<OfficePoint> contour) {
            cancellationToken.ThrowIfCancellationRequested();
            if (contour.Count > available - used) { exhausted = true; return false; }
            double left = double.PositiveInfinity, right = double.NegativeInfinity, top = double.PositiveInfinity, bottom = double.NegativeInfinity;
            foreach (OfficePoint point in contour) {
                left = Math.Min(left, point.X); right = Math.Max(right, point.X);
                top = Math.Min(top, point.Y); bottom = Math.Max(bottom, point.Y);
            }
            // A positive lexical width can still collapse at this coordinate's
            // floating-point precision. Fail closed instead of allocating invalid
            // paths or retrying unpaintable dots until the cursor stops advancing.
            if (!(right > left && bottom > top) || double.IsInfinity(left) || double.IsInfinity(right) ||
                double.IsInfinity(top) || double.IsInfinity(bottom)) { exhausted = true; return false; }
            used += contour.Count; contours.Add(contour); return true;
        }
        void Rectangle(double x, double length, double y) => Add(new[] {
            new OfficePoint(x, y - width / 2), new OfficePoint(x + length, y - width / 2),
            new OfficePoint(x + length, y + width / 2), new OfficePoint(x, y + width / 2)
        });
        void PaintRow(double y) {
            if (leader.Style == OfficeTextTabLineLeaderStyle.Solid) { Rectangle(0, gap, y); return; }
            if (leader.Style == OfficeTextTabLineLeaderStyle.Wave) { Wave(y); return; }
            // Pattern units are stroke widths. Dots are full circles; shortened final
            // dashes have butt ends so no outline can cover the following field.
            double[] pattern = leader.Style switch {
                OfficeTextTabLineLeaderStyle.Dotted => new[] { 1D, 3D },
                OfficeTextTabLineLeaderStyle.Dash => new[] { 4D, 3D },
                OfficeTextTabLineLeaderStyle.LongDash => new[] { 8D, 3D },
                OfficeTextTabLineLeaderStyle.DotDash => new[] { 1D, 3D, 4D, 3D },
                _ => new[] { 1D, 3D, 1D, 3D, 4D, 3D }
            };
            double x = 0; int index = 0;
            while (x < gap && !exhausted) {
                cancellationToken.ThrowIfCancellationRequested();
                double length = pattern[index] * width;
                if ((index & 1) == 0) {
                    if (pattern[index] == 1) {
                        if (length <= gap - x) Add(OfficeCurveFlattening.Ellipse(x + width / 2, y, width / 2, width / 2, 8));
                    } else Rectangle(x, Math.Min(length, gap - x), y);
                }
                double next = x + length;
                if (next <= x) { exhausted = true; break; }
                x = next; index = (index + 1) % pattern.Length;
            }
        }
        void Wave(double y) {
            // Leave half a stroke at both ends. Bevel joins keep every outline in
            // that gap even at a turning point and avoid miter overhangs.
            if (gap <= width) return;
            double start = width / 2, end = gap - width / 2;
            double wavelength = Math.Max(width * 8, fontSize * .4D), step = wavelength / 16;
            double requested = Math.Ceiling((end - start) / step);
            int count = (int)Math.Min(requested, Math.Max(0, (available - used) / 8));
            if (requested > count) exhausted = true;
            if (count < 1) return;
            var points = new List<OfficePoint>(count + 1) { new OfficePoint(start, y) };
            for (int i = 1; i <= count; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                double x = Math.Min(end, start + i * step);
                points.Add(new OfficePoint(x, y + amplitude * Math.Sin((x - start) / wavelength * Math.PI * 2)));
            }
            var pieces = new List<List<OfficePoint>>();
            OfficeRasterStroker.AppendOutline(points, width, OfficeStrokeLineCap.Butt, OfficeStrokeLineJoin.Bevel, 4, 8, pieces);
            foreach (var piece in pieces) if (!Add(piece)) break;
        }
    }
    private static bool FinitePositive(double value) => value > 0 && !double.IsNaN(value) && !double.IsInfinity(value);
}
