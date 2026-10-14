using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private static IReadOnlyList<List<OfficePoint>> TransformTextContours(IReadOnlyList<List<OfficePoint>> contours, double bottom, bool italic, double rotationRadians, double rotationCenterX, double rotationCenterY, bool flipHorizontal, bool flipVertical) {
        if ((!italic && Math.Abs(rotationRadians) < TextRotationEpsilon && !flipHorizontal && !flipVertical) || contours.Count == 0) {
            return contours;
        }

        List<List<OfficePoint>> transformed = new(contours.Count);
        foreach (List<OfficePoint> contour in contours) {
            List<OfficePoint> points = new(contour.Count);
            foreach (OfficePoint point in contour) {
                points.Add(TransformTextPoint(point, bottom, italic, rotationRadians, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical));
            }

            transformed.Add(points);
        }

        return transformed;
    }

    private static IReadOnlyList<List<OfficePoint>> TransformTextContours(IReadOnlyList<List<OfficePoint>> contours, double bottom, bool italic, OfficeTransform transform) {
        if ((!italic && transform == OfficeTransform.Identity) || contours.Count == 0) {
            return contours;
        }

        List<List<OfficePoint>> transformed = new(contours.Count);
        foreach (List<OfficePoint> contour in contours) {
            List<OfficePoint> points = new(contour.Count);
            foreach (OfficePoint point in contour) {
                OfficePoint skewed = italic ? new OfficePoint(point.X + ((bottom - point.Y) * ItalicShear), point.Y) : point;
                points.Add(transform.TransformPoint(skewed));
            }

            transformed.Add(points);
        }

        return transformed;
    }

    private static OfficePoint TransformTextPoint(OfficePoint point, double bottom, bool italic, double rotationRadians, double rotationCenterX, double rotationCenterY, bool flipHorizontal, bool flipVertical) {
        if (!italic && Math.Abs(rotationRadians) < TextRotationEpsilon && !flipHorizontal && !flipVertical) {
            return point;
        }

        OfficePoint skewed = italic ? new OfficePoint(point.X + ((bottom - point.Y) * ItalicShear), point.Y) : point;
        return TransformFramePoint(skewed, rotationRadians, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical);
    }

    private static OfficePoint TransformFramePoint(OfficePoint point, double rotationRadians, double centerX, double centerY, bool flipHorizontal, bool flipVertical) {
        OfficePoint transformed = point;
        if (flipHorizontal) {
            transformed = new OfficePoint((2D * centerX) - transformed.X, transformed.Y);
        }

        if (flipVertical) {
            transformed = new OfficePoint(transformed.X, (2D * centerY) - transformed.Y);
        }

        return Math.Abs(rotationRadians) < TextRotationEpsilon
            ? transformed
            : OfficeGeometry.RotatePoint(transformed, centerX, centerY, rotationRadians);
    }

    private static OfficePoint TransformAffineTextPoint(OfficePoint point, double bottom, bool italic, OfficeTransform transform) {
        OfficePoint skewed = italic ? new OfficePoint(point.X + ((bottom - point.Y) * ItalicShear), point.Y) : point;
        return transform.TransformPoint(skewed);
    }

    private static double GetAffineStrokeScale(OfficeTransform transform) {
        double xScale = Math.Sqrt((transform.M11 * transform.M11) + (transform.M12 * transform.M12));
        double yScale = Math.Sqrt((transform.M21 * transform.M21) + (transform.M22 * transform.M22));
        double scale = Math.Max(xScale, yScale);
        return !double.IsNaN(scale) && !double.IsInfinity(scale) && scale > 0D ? scale : 1D;
    }

    private static double ResolveRasterTextSize(double fontSize, double height, bool positioned) =>
        positioned ? Math.Max(0.1D, fontSize) : Math.Max(6D, Math.Min(fontSize, height - 2D));

    private static double ResolveRasterTextTop(IOfficeFontProgram font, double size, double height, double? baselineFontSize) {
        // Drawing text uses an alphabetic baseline one source em below its top,
        // matching the SVG positioned-text contract. Do not recenter it in the frame.
        if (baselineFontSize.HasValue) return baselineFontSize.Value - ResolveRasterBaseline(font, size);
        return Math.Max(1D, (height - font.LineHeight(size)) / 2D);
    }

    private static double ResolveRasterTextLineOutlineTop(IOfficeFontProgram font, double size, double top) =>
        top + size * .84D - ResolveRasterBaseline(font, size);

    private static double ResolveRasterBaseline(IOfficeFontProgram font, double size) {
        double lineHeight = font.LineHeight(size);
        double baseline = font is IOfficeFontBaselineMetrics metrics ? metrics.BaselineOffset(size) : lineHeight * 0.8D;
        return double.IsNaN(baseline) || double.IsInfinity(baseline)
            ? lineHeight * 0.8D : Math.Max(0D, Math.Min(lineHeight, baseline));
    }
}
