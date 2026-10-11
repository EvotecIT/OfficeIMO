using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Measures projected fill and stroke after the shape's actual clip geometry.</summary>
    private static bool TryGetShapePaintBounds(OfficeShape shape, double x, double bottomY,
        PdfOptions? options, string source, string contentKind,
        out double left, out double bottom, out double width, out double height) {
        try {
            return TryMeasureShapePaintBounds(shape, x, bottomY, options, source, contentKind,
                out left, out bottom, out width, out height);
        } catch (Exception error) when (error is InvalidOperationException or NotSupportedException) {
            // Diagnostic inspection may exhaust its bounded geometry work even
            // when the native writer can emit the original path and dash array.
            options?.AddUnassessedLayoutWarning("HeaderFooterPaintBoundsUnmeasured", source,
                source + " " + contentKind + " paint bounds exceed the geometry inspection limit; physical page clipping is unassessed.");
            left = bottom = width = height = 0D;
            return false;
        }
    }

    private static bool TryMeasureShapePaintBounds(OfficeShape shape, double x, double bottomY,
        PdfOptions? options, string source, string contentKind,
        out double left, out double bottom, out double width, out double height) {
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
            var effectiveClips = clips == null ? new List<OfficeTextInkClip>() : new List<OfficeTextInkClip>(clips);
            var linear = stroke ? geometry.StrokeGradient : geometry.FillGradient;
            var radial = stroke ? geometry.StrokeRadialGradient : geometry.FillRadialGradient;
            if (!TryCreateGradientAlphaContours(geometry, paint, linear, radial, out var alpha)) {
                options?.AddUnassessedLayoutWarning("HeaderFooterPaintBoundsUnmeasured", source,
                    source + " " + contentKind + " gradient alpha bounds are unassessed; physical page clipping is unassessed.");
                return;
            }
            if (alpha != null) {
                if (alpha.Count == 0) return;
                var mappedAlpha = new List<List<OfficePoint>>();
                foreach (var ring in alpha) mappedAlpha.Add(Map(ring, geometry));
                effectiveClips.Add(OfficeTextInkClip.FromFilledContours(mappedAlpha, OfficeFillRule.EvenOdd));
            }
            var mapped = new List<List<OfficePoint>>();
            foreach (var ring in paint) mapped.Add(Map(ring, geometry));
            var bounds = OfficeRasterCanvas.MeasureNominalFilledContourBounds(mapped, rule, effectiveClips);
            if (!bounds.IsMeasured) {
                options?.AddUnassessedLayoutWarning("HeaderFooterPaintBoundsUnmeasured", source,
                    source + " " + contentKind + " paint bounds exceed the contour inspection limit; physical page clipping is unassessed.");
            } else if (bounds.HasInk) {
                IncludeShapeBounds(bounds.Left, bounds.Top, bounds.Right - bounds.Left, bounds.Bottom - bounds.Top,
                    ref hasBounds, ref paintLeft, ref paintBottom, ref paintWidth, ref paintHeight);
            }
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
