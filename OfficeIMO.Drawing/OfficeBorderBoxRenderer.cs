using System;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>
/// Renders rectangular border boxes through the shared raster and SVG primitives.
/// </summary>
public static class OfficeBorderBoxRenderer {
    /// <summary>
    /// Draws a border box onto a raster canvas using line-on-bounds geometry.
    /// </summary>
    public static void DrawRaster(OfficeRasterCanvas canvas, double x, double y, double width, double height, OfficeBorderBox borders) {
        if (canvas == null) {
            throw new ArgumentNullException(nameof(canvas));
        }

        ValidateBounds(x, y, width, height);
        DrawRasterLine(canvas, x, y, x + width, y, borders.Top);
        DrawRasterLine(canvas, x + width, y, x + width, y + height, borders.Right);
        DrawRasterLine(canvas, x, y + height, x + width, y + height, borders.Bottom);
        DrawRasterLine(canvas, x, y, x, y + height, borders.Left);
        DrawRasterLine(canvas, x, y, x + width, y + height, borders.DiagonalDown);
        DrawRasterLine(canvas, x, y + height, x + width, y, borders.DiagonalUp);
    }

    /// <summary>
    /// Appends SVG line elements for a border box using line-on-bounds geometry.
    /// </summary>
    public static void AppendSvg(StringBuilder builder, double x, double y, double width, double height, OfficeBorderBox borders) {
        if (builder == null) {
            throw new ArgumentNullException(nameof(builder));
        }

        ValidateBounds(x, y, width, height);
        AppendSvgLine(builder, x, y, x + width, y, borders.Top);
        AppendSvgLine(builder, x + width, y, x + width, y + height, borders.Right);
        AppendSvgLine(builder, x, y + height, x + width, y + height, borders.Bottom);
        AppendSvgLine(builder, x, y, x, y + height, borders.Left);
        AppendSvgLine(builder, x, y, x + width, y + height, borders.DiagonalDown);
        AppendSvgLine(builder, x, y + height, x + width, y, borders.DiagonalUp);
    }

    private static void DrawRasterLine(OfficeRasterCanvas canvas, double x1, double y1, double x2, double y2, OfficeBorderSide? side) {
        if (!side.HasValue || !side.Value.IsVisible) {
            return;
        }

        OfficeBorderSide borderSide = side.Value;
        if (borderSide.LineKind == OfficeBorderLineKind.Double) {
            canvas.DrawParallelStyledLine(x1, y1, x2, y2, borderSide.Color, borderSide.Width, ResolveDoubleLineSeparation(borderSide));
            return;
        }

        canvas.DrawStyledLine(x1, y1, x2, y2, borderSide.Color, borderSide.Width, borderSide.DashStyle);
    }

    private static void AppendSvgLine(StringBuilder builder, double x1, double y1, double x2, double y2, OfficeBorderSide? side) {
        if (!side.HasValue || !side.Value.IsVisible) {
            return;
        }

        OfficeBorderSide borderSide = side.Value;
        if (borderSide.LineKind == OfficeBorderLineKind.Double) {
            builder.AppendParallelLineElements(x1, y1, x2, y2, borderSide.Color, borderSide.Width, ResolveDoubleLineSeparation(borderSide), borderSide.DashStyle, ResolveLineCap(borderSide));
            return;
        }

        builder.AppendLineElement(x1, y1, x2, y2, borderSide.Color, borderSide.Width, borderSide.DashStyle, ResolveLineCap(borderSide));
    }

    private static OfficeStrokeLineCap? ResolveLineCap(OfficeBorderSide side) =>
        side.DashStyle == OfficeStrokeDashStyle.Dot
            ? OfficeStrokeLineCap.Round
            : null;

    private static double ResolveDoubleLineSeparation(OfficeBorderSide side) =>
        side.DoubleLineSeparation > 0D ? side.DoubleLineSeparation : Math.Max(1D, side.Width * 3D);

    private static void ValidateBounds(double x, double y, double width, double height) {
        ValidateFiniteNonNegative(x, nameof(x));
        ValidateFiniteNonNegative(y, nameof(y));
        ValidatePositiveFinite(width, nameof(width));
        ValidatePositiveFinite(height, nameof(height));
    }

    private static void ValidateFiniteNonNegative(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D) {
            throw new ArgumentOutOfRangeException(paramName, "Border box coordinates must be finite non-negative numbers.");
        }
    }

    private static void ValidatePositiveFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0D) {
            throw new ArgumentOutOfRangeException(paramName, "Border box dimensions must be finite positive numbers.");
        }
    }
}
