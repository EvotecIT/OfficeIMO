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

        ResolvedDataBar bar = Resolve(x, y, width, height, startRatio, ratio, verticalInset);
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

        ResolvedDataBar bar = Resolve(x, y, width, height, startRatio, ratio, verticalInset);
        if (bar.Width <= 0D) {
            return builder;
        }

        var attributes = new StringBuilder();
        attributes.AppendPaintAttribute("fill", color);
        builder.AppendRectElement(bar.X, bar.Y, bar.Width, bar.Height, attributes.ToString());
        return builder;
    }

    private static ResolvedDataBar Resolve(double x, double y, double width, double height, double startRatio, double ratio, double verticalInset) {
        double inset = Math.Max(0D, verticalInset);
        double barX = x + (width * ClampUnit(startRatio));
        double barY = y + inset;
        double barWidth = Math.Max(0D, width * ClampUnit(ratio));
        double barHeight = Math.Max(1D, height - (inset * 2D));
        return new ResolvedDataBar(barX, barY, barWidth, barHeight);
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

    private readonly struct ResolvedDataBar {
        internal ResolvedDataBar(double x, double y, double width, double height) {
            X = x;
            Y = y;
            Width = width;
            Height = height;
        }

        internal double X { get; }

        internal double Y { get; }

        internal double Width { get; }

        internal double Height { get; }
    }
}
