using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    private static void AddBubbleMarker(OfficeDrawing drawing, OfficeChartSeries series,
        int pointIndex, OfficePoint center, double maximumBubbleSize,
        double maximumDiameter, OfficeChartBubbleSizeMode sizeMode,
        OfficeColor color) {
        if (series.BubbleSizes == null || pointIndex < 0 ||
            pointIndex >= series.BubbleSizes.Count) {
            return;
        }

        double size = series.BubbleSizes[pointIndex];
        if (size <= 0D) {
            return;
        }

        double ratio = Math.Min(1D, size / maximumBubbleSize);
        double sizeFactor = sizeMode == OfficeChartBubbleSizeMode.Width
            ? ratio
            : Math.Sqrt(ratio);
        double minimumDiameter = Math.Min(3D, maximumDiameter);
        double diameter = maximumBubbleSize <= 0D
            ? 0D
            : minimumDiameter + (maximumDiameter - minimumDiameter) * sizeFactor;
        if (diameter <= 0D) {
            return;
        }
        double outlineWidth = series.MarkerOutlineWidth ?? 1D;
        OfficeColor outlineColor = series.MarkerOutlineColor ?? color;
        AddShape(drawing, OfficeShape.Ellipse(diameter, diameter),
            center.X - diameter / 2D, center.Y - diameter / 2D,
            color, outlineColor, outlineWidth);
    }

    private static double GetMaximumBubbleDiameter(double plotWidth,
        double plotHeight, double bubbleScalePercent) {
        if (bubbleScalePercent <= 0D) {
            return 0D;
        }

        double shortestSide = Math.Min(plotWidth, plotHeight);
        double defaultDiameter = Math.Max(12D,
            Math.Min(42D, shortestSide * 0.16D));
        return Math.Min(shortestSide * 0.8D,
            defaultDiameter * bubbleScalePercent / 100D);
    }

    private static double GetMaximumBubbleSize(
        System.Collections.Generic.IReadOnlyList<OfficeChartSeries> series) {
        double maximum = 0D;
        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++) {
            IReadOnlyList<double>? sizes = series[seriesIndex].BubbleSizes;
            if (sizes == null) continue;
            for (int pointIndex = 0; pointIndex < sizes.Count; pointIndex++) {
                maximum = Math.Max(maximum, sizes[pointIndex]);
            }
        }
        return maximum;
    }
}
