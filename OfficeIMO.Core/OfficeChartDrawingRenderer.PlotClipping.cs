using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    private readonly struct ChartPlotBounds {
        internal ChartPlotBounds(double left, double top, double width, double height) {
            Left = left;
            Top = top;
            Width = width;
            Height = height;
        }

        internal double Left { get; }
        internal double Top { get; }
        internal double Width { get; }
        internal double Height { get; }

        internal bool Contains(OfficePoint point) =>
            point.X >= Left && point.X <= Left + Width &&
            point.Y >= Top && point.Y <= Top + Height;
    }

    private static bool HasExplicitValueBounds(OfficeChartLayout layout, OfficeChartAxisGroup? axisGroup) =>
        axisGroup == OfficeChartAxisGroup.Secondary
            ? layout.SecondaryValueAxis?.Minimum.HasValue == true || layout.SecondaryValueAxis?.Maximum.HasValue == true
            : layout.VerticalAxisMinimum.HasValue || layout.VerticalAxisMaximum.HasValue;

    private static double ToUnclampedPlotY(double value, double min, double max, double plotTop, double plotHeight) =>
        plotTop + plotHeight * (1D - BoundedPlotRatio(value, min, max));

    private static double ToUnclampedPlotX(double value, double min, double max, double plotLeft, double plotWidth) =>
        plotLeft + plotWidth * BoundedPlotRatio(value, min, max);

    private static double BoundedPlotRatio(double value, double min, double max) {
        double range = max - min;
        if (range <= 0D || double.IsInfinity(range)) return 0.5D;
        double ratio = (value - min) / range;
        if (double.IsNaN(ratio)) return 0.5D;
        return Math.Max(-1000000D, Math.Min(1000000D, ratio));
    }

    private static void AddClippedPlotGeometry(OfficeDrawing drawing, OfficeDrawing geometry,
        OfficeDrawing labels, ChartPlotBounds bounds) {
        drawing.AddClippedDrawing(geometry, bounds.Left, bounds.Top,
            OfficeClipPath.Rectangle(bounds.Width, bounds.Height), -bounds.Left, -bounds.Top);
        drawing.AddDrawing(labels, 0D, 0D);
    }
}
