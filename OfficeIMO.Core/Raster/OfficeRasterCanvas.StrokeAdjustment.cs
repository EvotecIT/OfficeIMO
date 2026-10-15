using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private const double AxisAlignmentTolerance = 1e-6D;

    /// <summary>
    /// PDF-style stroke adjustment: axis-aligned strokes and rectangle fills snap to whole device pixels
    /// and every stroke paints at least one device pixel (ISO 32000-1, 10.7.5), as common PDF viewers do.
    /// </summary>
    internal bool StrokeAdjustment { get; set; }

    /// <summary>Enables stroke adjustment, whole-pixel text baselines and PDF image sampling together.</summary>
    internal bool PdfRasterConventions {
        get => StrokeAdjustment && SnapTextBaselines && PdfImageSampling && AntialiasedClipping;
        set => StrokeAdjustment = SnapTextBaselines = PdfImageSampling = AntialiasedClipping = value;
    }

    /// <summary>Places upright positioned-text baselines on whole device rows (the glyph origin moves up to the row).</summary>
    internal bool SnapTextBaselines { get; set; }

    /// <summary>Returns the logical top offset that moves a baseline onto the device row at or above it.</summary>
    private double SnapBaselineOffset(double logicalBaseline) {
        if (!SnapTextBaselines || CoordinateScaleY <= 0D || !IsFinite(logicalBaseline)) return 0D;
        double device = logicalBaseline * CoordinateScaleY;
        return (Math.Floor(device + 1e-6D) - device) / CoordinateScaleY;
    }

    private (OfficeColor Color, double Width, bool FillSuppressed)? _textOutlineStroke;

    /// <summary>Strokes the outlines of glyphs drawn within the scope, optionally without filling them.</summary>
    internal IDisposable PushTextOutlineStroke(OfficeColor color, double width, bool fillSuppressed) {
        var previous = _textOutlineStroke;
        _textOutlineStroke = (color, width, fillSuppressed);
        return new TextOutlineStrokeScope(this, previous);
    }

    private sealed class TextOutlineStrokeScope : IDisposable {
        private OfficeRasterCanvas? _canvas;
        private readonly (OfficeColor Color, double Width, bool FillSuppressed)? _previous;
        internal TextOutlineStrokeScope(OfficeRasterCanvas canvas, (OfficeColor Color, double Width, bool FillSuppressed)? previous) { _canvas = canvas; _previous = previous; }
        public void Dispose() { if (_canvas != null) { _canvas._textOutlineStroke = _previous; _canvas = null; } }
    }

    // Glyph outlines are closed contours stroked with the PDF text-state defaults (miter joins, limit 10).
    private void PaintTextContours(List<List<OfficePoint>> contours, OfficeColor fill, double emboldenOffset) {
        if (_textOutlineStroke is not { } stroke) {
            FillTextContours(contours, fill, emboldenOffset);
            return;
        }
        if (!stroke.FillSuppressed) FillTextContours(contours, fill, emboldenOffset);
        var outlines = new List<OfficeFlattenedPathContour>(contours.Count);
        foreach (List<OfficePoint> contour in contours) if (contour.Count >= 2) outlines.Add(new OfficeFlattenedPathContour(contour, true));
        if (outlines.Count == 0) return;
        OfficeColor strokeColor = stroke.Color;
        StrokeContours(outlines, stroke.Width > 0D ? stroke.Width : 1E-9D, OfficeStrokeLineCap.Butt, OfficeStrokeLineJoin.Miter, 10D,
            null, 0D, (_, _) => strokeColor, hairline: stroke.Width <= 0D);
    }

    private double DevicePixelsPerLogicalUnit => Math.Sqrt(Math.Abs(CoordinateScaleX * CoordinateScaleY));

    /// <summary>Returns the logical stroke width after hairline and minimum-width adjustment.</summary>
    private double AdjustStrokeWidth(double width, double devicePixelsPerUnit, bool hairline) {
        if (devicePixelsPerUnit <= 0D || !IsFinite(devicePixelsPerUnit)) return width;
        double onePixel = 1D / devicePixelsPerUnit;
        if (hairline) return onePixel;
        return StrokeAdjustment && width < onePixel ? onePixel : width;
    }

    /// <summary>Snaps a logical rectangle fill to device pixels, keeping at least one pixel on each axis.</summary>
    internal bool TrySnapRectangle(ref double x, ref double y, ref double width, ref double height) {
        if (!StrokeAdjustment || !IsFinite(x) || !IsFinite(y) || !IsFinite(width) || !IsFinite(height)) return false;
        double sx = CoordinateScaleX, sy = CoordinateScaleY;
        if (sx <= 0D || sy <= 0D) return false;
        SnapSpan(x * sx, (x + width) * sx, out double left, out double right);
        SnapSpan(y * sy, (y + height) * sy, out double top, out double bottom);
        x = left / sx; width = (right - left) / sx;
        y = top / sy; height = (bottom - top) / sy;
        return true;
    }

    /// <summary>
    /// Fills axis-aligned solid strokes as unions of pixel-snapped rectangles. Returns false when any contour
    /// is not axis-aligned in device space, so the caller uses the general stroker.
    /// </summary>
    private bool TryStrokeAdjustedContours(IReadOnlyList<OfficeFlattenedPathContour> contours, double width,
        OfficeStrokeLineCap cap, IReadOnlyList<double>? pattern, OfficeTransform transform,
        Func<double, double, OfficeColor> paint) {
        if (!StrokeAdjustment || pattern is { Count: > 0 } || width <= 0D || !IsFinite(width)) return false;
        double sx = CoordinateScaleX, sy = CoordinateScaleY;
        if (sx <= 0D || sy <= 0D) return false;
        var rectangles = new List<(double Left, double Top, double Right, double Bottom)>();
        foreach (OfficeFlattenedPathContour contour in contours) {
            IReadOnlyList<OfficePoint> points = contour.Points;
            int count = points.Count;
            if (contour.Closed && count > 1 && points[0].X == points[count - 1].X && points[0].Y == points[count - 1].Y) count--;
            int segments = contour.Closed ? count : count - 1;
            for (int index = 0; index < segments; index++) {
                OfficePoint start = points[index];
                OfficePoint end = points[(index + 1) % count];
                double dx = end.X - start.X, dy = end.Y - start.Y;
                double length = Math.Sqrt(dx * dx + dy * dy);
                if (length <= 0D) continue;
                double ux = dx / length * width / 2D, uy = dy / length * width / 2D;
                // Device offsets of the half-width normal and of the half-width extension along the segment.
                OfficePoint normal = Linear(transform, -uy, ux, sx, sy);
                OfficePoint along = Linear(transform, ux, uy, sx, sy);
                OfficePoint a = Device(transform, start, sx, sy);
                OfficePoint b = Device(transform, end, sx, sy);
                bool horizontal = Math.Abs(b.Y - a.Y) <= AxisAlignmentTolerance * Math.Max(1D, Math.Abs(b.X - a.X));
                bool vertical = Math.Abs(b.X - a.X) <= AxisAlignmentTolerance * Math.Max(1D, Math.Abs(b.Y - a.Y));
                if (horizontal == vertical) return false;
                double normalSkew = horizontal ? Math.Abs(normal.X) : Math.Abs(normal.Y);
                if (normalSkew > AxisAlignmentTolerance * Math.Max(1D, Math.Abs(horizontal ? normal.Y : normal.X))) return false;
                bool extendStart = contour.Closed || index > 0 || cap != OfficeStrokeLineCap.Butt && contour.UseStartLineCap;
                bool extendEnd = contour.Closed || index < segments - 1 || cap != OfficeStrokeLineCap.Butt && contour.UseEndLineCap;
                double ax = a.X - (extendStart ? along.X : 0D), ay = a.Y - (extendStart ? along.Y : 0D);
                double bx = b.X + (extendEnd ? along.X : 0D), by = b.Y + (extendEnd ? along.Y : 0D);
                double halfX = Math.Abs(normal.X), halfY = Math.Abs(normal.Y);
                double left = Math.Min(ax, bx) - halfX, right = Math.Max(ax, bx) + halfX;
                double top = Math.Min(ay, by) - halfY, bottom = Math.Max(ay, by) + halfY;
                SnapSpan(left, right, out left, out right);
                SnapSpan(top, bottom, out top, out bottom);
                rectangles.Add((left, top, right, bottom));
            }
        }
        if (rectangles.Count == 0) return true;
        var fill = new List<IReadOnlyList<OfficePoint>>(rectangles.Count);
        foreach (var rectangle in rectangles) {
            fill.Add(new[] {
                new OfficePoint(rectangle.Left / sx, rectangle.Top / sy),
                new OfficePoint(rectangle.Right / sx, rectangle.Top / sy),
                new OfficePoint(rectangle.Right / sx, rectangle.Bottom / sy),
                new OfficePoint(rectangle.Left / sx, rectangle.Bottom / sy)
            });
        }
        FillContourPaint(fill, OfficeFillRule.NonZero, paint);
        return true;
    }

    // Edges round to the nearest pixel boundary; a span thinner than one pixel keeps the pixel holding its centre.
    private static void SnapSpan(double low, double high, out double snappedLow, out double snappedHigh) {
        if (high < low) (low, high) = (high, low);
        snappedLow = Math.Floor(low + 0.5D);
        snappedHigh = Math.Floor(high + 0.5D);
        if (snappedHigh <= snappedLow) {
            snappedLow = Math.Floor((low + high) / 2D);
            snappedHigh = snappedLow + 1D;
        }
    }

    private static OfficePoint Device(OfficeTransform transform, OfficePoint point, double sx, double sy) {
        OfficePoint logical = transform.TransformPoint(point);
        return new OfficePoint(logical.X * sx, logical.Y * sy);
    }

    private static OfficePoint Linear(OfficeTransform transform, double x, double y, double sx, double sy) =>
        new OfficePoint((x * transform.M11 + y * transform.M21) * sx, (x * transform.M12 + y * transform.M22) * sy);
}
