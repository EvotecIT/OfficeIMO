using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Measures projected fill and stroke after the shape's actual clip geometry.</summary>
    private static bool TryGetShapePaintBounds(OfficeShape shape, double x, double bottomY,
        PdfOptions? options, string source, string contentKind,
        out double left, out double bottom, out double width, out double height, bool conservativeWhenUnmeasured = false) {
        try {
            return TryMeasureShapePaintBounds(shape, x, bottomY, options, source, contentKind,
                out left, out bottom, out width, out height, conservativeWhenUnmeasured);
        } catch (Exception error) when (error is InvalidOperationException or NotSupportedException) {
            // Diagnostic inspection may exhaust its bounded geometry work even
            // when the native writer can emit the original path and dash array.
            options?.AddUnassessedLayoutWarning("HeaderFooterPaintBoundsUnmeasured", source,
                source + " " + contentKind + " paint bounds exceed the geometry inspection limit; physical page clipping is unassessed.");
            left = bottom = width = height = 0D;
            if (!conservativeWhenUnmeasured) return false;
            var bounds = shape.Kind == OfficeShapeKind.Path ? OfficePathGeometry.Bounds(shape.PathCommands)
                : (Left: 0D, Top: 0D, Right: shape.Width, Bottom: shape.Height);
            double outset = shape.StrokeWidth > 0 && HasVisibleBoundsPaint(shape.StrokeColor, shape.StrokeGradient, shape.StrokeRadialGradient, shape.StrokeOpacity)
                ? shape.StrokeWidth / 2 * Math.Max(1D, PdfPaintMiterLimit(shape)) : 0D;
            var mapped = (shape.Transform ?? OfficeTransform.Identity).TransformRectangleBounds(
                bounds.Left - outset, bounds.Top - outset, bounds.Right - bounds.Left + 2 * outset, bounds.Bottom - bounds.Top + 2 * outset);
            left = x + mapped.Left; bottom = bottomY + shape.Height - mapped.Bottom;
            width = mapped.Right - mapped.Left; height = mapped.Bottom - mapped.Top;
            return width > 0 && height > 0;
        }
    }

    private static bool TryMeasureShapePaintBounds(OfficeShape shape, double x, double bottomY,
        PdfOptions? options, string source, string contentKind,
        out double left, out double bottom, out double width, out double height, bool conservativeWhenUnmeasured) {
        left = bottom = width = height = 0D;
        if (shape.ClipPath?.Kind == OfficeClipPathKind.Empty) return false;
        IReadOnlyList<OfficeTextInkClip>? clips = shape.ClipPath == null ? null : new[] {
            OfficeTextInkClip.FromFilledContours(OfficeClipPathGeometry.CreateContours(shape.ClipPath, points => Map(points, shape)), shape.ClipPath.FillRule)
        };
        bool hasBounds = false;
        double paintLeft = 0D, paintBottom = 0D, paintWidth = 0D, paintHeight = 0D;
        if (shape.Kind != OfficeShapeKind.Line && HasVisibleBoundsPaint(shape.FillColor, shape.FillGradient, shape.FillRadialGradient, shape.FillOpacity))
            AddFill(shape);
        if (shape.StrokeWidth > 0D && HasVisibleBoundsPaint(shape.StrokeColor, shape.StrokeGradient, shape.StrokeRadialGradient, shape.StrokeOpacity)) {
            if (shape.StrokeGradient != null || shape.StrokeRadialGradient != null) {
                // Reuse the painter's outline, including its cap/join defaults.
                OfficeShape? projected = CreateGradientStrokeShape(shape);
                if (projected != null) AddFill(projected);
            } else {
                var paint = new List<List<OfficePoint>>();
                double density = OfficeStrokeGeometry.LocalFlatteningDensity(shape.Transform, 4D);
                var contours = OfficeStrokeGeometry.FlattenShape(shape, density);
                bool customDash = shape.StrokeDashArray.Count > 0;
                IReadOnlyList<double>? dash = customDash ? shape.StrokeDashArray : PdfPresetDashPattern(shape.StrokeDashStyle, shape.StrokeWidth);
                foreach (var ring in OfficeStrokeGeometry.Create(contours, shape.StrokeWidth,
                    shape.StrokeLineCap ?? (customDash ? OfficeStrokeLineCap.Butt : PdfPresetLineCap(shape.StrokeDashStyle)), shape.StrokeLineJoin ?? OfficeStrokeLineJoin.Miter,
                    PdfPaintMiterLimit(shape), dash, shape.StrokeDashOffset, density,
                    double.NegativeInfinity, double.NegativeInfinity, double.PositiveInfinity, double.PositiveInfinity))
                    paint.Add(ring);
                Measure(paint, OfficeFillRule.NonZero, shape, true);
            }
        }
        left = paintLeft; bottom = paintBottom; width = paintWidth; height = paintHeight;
        return hasBounds;

        void AddFill(OfficeShape geometry) {
            var paint = new List<List<OfficePoint>>();
            foreach (var contour in OfficeStrokeGeometry.FlattenShape(geometry,
                OfficeStrokeGeometry.LocalFlatteningDensity(geometry.Transform, 4D))) paint.Add(new List<OfficePoint>(contour.Points));
            Measure(paint, geometry.FillRule, geometry, false);
        }
        void Measure(List<List<OfficePoint>> paint, OfficeFillRule rule, OfficeShape geometry, bool stroke) {
            List<List<OfficePoint>>? mapped = null;
            List<List<OfficePoint>> MappedPaint() {
                if (mapped != null) return mapped;
                mapped = new List<List<OfficePoint>>();
                foreach (var ring in paint) mapped.Add(Map(ring, geometry));
                return mapped;
            }
            var effectiveClips = clips == null ? new List<OfficeTextInkClip>() : new List<OfficeTextInkClip>(clips);
            var linear = stroke ? geometry.StrokeGradient : geometry.FillGradient;
            var radial = stroke ? geometry.StrokeRadialGradient : geometry.FillRadialGradient;
            if (!TryCreateGradientAlphaContours(geometry, paint, linear, radial, out var alpha)) {
                options?.AddUnassessedLayoutWarning("HeaderFooterPaintBoundsUnmeasured", source,
                    source + " " + contentKind + " gradient alpha bounds are unassessed; physical page clipping is unassessed.");
                if (conservativeWhenUnmeasured) IncludeUnmeasuredGeometry(MappedPaint());
                return;
            }
            if (alpha != null) {
                if (alpha.Count == 0) return;
                var mappedAlpha = new List<List<OfficePoint>>();
                foreach (var ring in alpha) mappedAlpha.Add(Map(ring, geometry));
                effectiveClips.Add(OfficeTextInkClip.FromFilledContours(mappedAlpha, OfficeFillRule.EvenOdd));
            }
            var bounds = OfficeRasterCanvas.MeasureNominalFilledContourBounds(MappedPaint(), rule, effectiveClips);
            if (!bounds.IsMeasured) {
                options?.AddUnassessedLayoutWarning("HeaderFooterPaintBoundsUnmeasured", source,
                    source + " " + contentKind + " paint bounds exceed the contour inspection limit; physical page clipping is unassessed.");
                if (conservativeWhenUnmeasured) IncludeUnmeasuredGeometry(MappedPaint());
            } else if (bounds.HasInk) {
                IncludeShapeBounds(bounds.Left, bounds.Top, bounds.Right - bounds.Left, bounds.Bottom - bounds.Top,
                    ref hasBounds, ref paintLeft, ref paintBottom, ref paintWidth, ref paintHeight);
            }
        }
        void IncludeUnmeasuredGeometry(List<List<OfficePoint>> mapped) {
            double minX = double.PositiveInfinity, minY = double.PositiveInfinity;
            double maxX = double.NegativeInfinity, maxY = double.NegativeInfinity;
            foreach (var ring in mapped) foreach (var point in ring) {
                minX = Math.Min(minX, point.X); minY = Math.Min(minY, point.Y);
                maxX = Math.Max(maxX, point.X); maxY = Math.Max(maxY, point.Y);
            }
            if (maxX > minX && maxY > minY) IncludeShapeBounds(minX, minY, maxX - minX, maxY - minY,
                ref hasBounds, ref paintLeft, ref paintBottom, ref paintWidth, ref paintHeight);
        }
        List<OfficePoint> Map(IReadOnlyList<OfficePoint> ring, OfficeShape geometry) {
            var mapped = new List<OfficePoint>(ring.Count);
            OfficeTransform transform = geometry.Transform ?? OfficeTransform.Identity;
            foreach (OfficePoint point in ring) {
                OfficePoint transformed = transform.TransformPoint(point);
                mapped.Add(new OfficePoint(x + transformed.X, bottomY + geometry.Height - transformed.Y));
            }
            return mapped;
        }
    }

    private static bool HasVisibleBoundsPaint(OfficeColor? color, OfficeLinearGradient? linear, OfficeRadialGradient? radial, double? opacity) {
        if ((opacity ?? 1D) <= 0D) return false;
        if (radial != null) {
            if (radial.OutsideColor?.A > 0) return true;
            foreach (OfficeGradientStop stop in radial.Stops) if (stop.Color.A > 0) return true;
            return false;
        }
        if (linear != null) {
            foreach (OfficeGradientStop stop in linear.Stops) if (stop.Color.A > 0) return true;
            return false;
        }
        return color.HasValue && color.Value.A > 0;
    }
}
