using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    private static void AddBubbleMarker(OfficeDrawing drawing, OfficeChartSeries series,
        int pointIndex, OfficePoint center, double plotWidth, double plotHeight,
        double maximumBubbleSize, OfficeColor color) {
        if (series.BubbleSizes == null || pointIndex < 0 ||
            pointIndex >= series.BubbleSizes.Count) {
            return;
        }

        double size = series.BubbleSizes[pointIndex];
        if (size <= 0D) {
            return;
        }

        double maximumDiameter = Math.Max(12D, Math.Min(42D, Math.Min(plotWidth, plotHeight) * 0.16D));
        double diameter = maximumBubbleSize <= 0D
            ? 0D
            : 3D + (maximumDiameter - 3D) * Math.Sqrt(size / maximumBubbleSize);
        double outlineWidth = series.MarkerOutlineWidth ?? 1D;
        OfficeColor outlineColor = series.MarkerOutlineColor ?? color;
        AddShape(drawing, OfficeShape.Ellipse(diameter, diameter),
            center.X - diameter / 2D, center.Y - diameter / 2D,
            color, outlineColor, outlineWidth);
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
