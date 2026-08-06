using System;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared dependency-free renderer for proportional data bars inside a rectangular region.
/// </summary>
public static class OfficeDataBarRenderer {
    /// <summary>
    /// Draws a resolved data bar on a raster canvas.
    /// </summary>
    public static void DrawRaster(
        OfficeRasterCanvas canvas,
        double x,
        double y,
        double width,
        double height,
        double startRatio,
        double ratio,
        OfficeColor color,
        double verticalInset = 2D) {
        if (canvas == null) {
            throw new ArgumentNullException(nameof(canvas));
        }

        OfficeDataBarGeometry bar = Resolve(x, y, width, height, startRatio, ratio, verticalInset);
        if (bar.Width <= 0D) {
            return;
        }

        canvas.FillRectangle(bar.X, bar.Y, bar.Width, bar.Height, color);
    }

    /// <summary>
    /// Appends SVG markup for a resolved data bar.
    /// </summary>
    public static StringBuilder AppendSvg(
        StringBuilder builder,
        double x,
        double y,
        double width,
        double height,
        double startRatio,
        double ratio,
        OfficeColor color,
        double verticalInset = 2D) {
        if (builder == null) {
            throw new ArgumentNullException(nameof(builder));
        }

        OfficeDataBarGeometry bar = Resolve(x, y, width, height, startRatio, ratio, verticalInset);
        if (bar.Width <= 0D) {
            return builder;
        }

        var attributes = new StringBuilder();
        attributes.AppendPaintAttribute("fill", color);
        builder.AppendRectElement(bar.X, bar.Y, bar.Width, bar.Height, attributes.ToString());
        return builder;
    }

    /// <summary>
    /// Resolves data-bar placement and size without binding the result to a specific output format.
    /// </summary>
    public static OfficeDataBarGeometry Resolve(
        double x,
        double y,
        double width,
        double height,
        double startRatio,
        double ratio,
        double verticalInset = 2D,
        double minimumHeight = 1D) {
        double inset = Math.Max(0D, verticalInset);
        double barX = x + (width * ClampUnit(startRatio));
        double barY = y + inset;
        double barWidth = Math.Max(0D, width * ClampUnit(ratio));
        double minHeight = Math.Max(0D, minimumHeight);
        double barHeight = Math.Max(minHeight, height - (inset * 2D));
        return new OfficeDataBarGeometry(barX, barY, barWidth, barHeight);
    }

    private static double ClampUnit(double value) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            return 0D;
        }

        if (value < 0D) {
            return 0D;
        }

        return value > 1D ? 1D : value;
    }
}
