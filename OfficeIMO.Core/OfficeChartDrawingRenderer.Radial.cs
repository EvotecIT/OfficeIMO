using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    private static void GetRadialPlotGeometry(double width, double height, double topBand,
        double bottomBand, double sideBand, bool leftLegend, double horizontalPadding,
        double verticalPadding, out double centerX, out double centerY, out double radius,
        out double visualWidth, out double contentHeight) {
        // Legend and title bands must leave a plot inside the authored frame, even on tiny exports.
        sideBand = Math.Max(0D, Math.Min(sideBand, width - Math.Min(1D, width)));
        topBand = Math.Max(0D, Math.Min(topBand, height - Math.Min(1D, height)));
        bottomBand = Math.Max(0D, Math.Min(bottomBand, height - topBand - Math.Min(1D, height)));
        visualWidth = width - sideBand;
        contentHeight = height - topBand - bottomBand;
        centerX = (leftLegend ? sideBand : 0D) + visualWidth / 2D;
        centerY = topBand + contentHeight / 2D;
        double preferredRadius = Math.Max(28D,
            Math.Min(visualWidth - horizontalPadding, contentHeight - verticalPadding) / 2D);
        radius = Math.Min(preferredRadius, Math.Min(visualWidth, contentHeight) / 2D);
    }

    private static void AddPieSeries(OfficeDrawing drawing, OfficeChartSnapshot snapshot, double width, double height, double contentTop, double bottomLegendHeight, bool doughnut, OfficeChartStyle style, OfficeChartLayout layout) {
        IReadOnlyList<string> categories = snapshot.Data.Categories;
        IReadOnlyList<OfficeChartSeries> series = snapshot.Data.Series;
        if (categories.Count == 0 || series.Count == 0) {
            return;
        }

        if (doughnut) {
            AddDoughnutSeries(drawing, snapshot, width, height, contentTop, bottomLegendHeight, style, layout);
            return;
        }

        OfficeChartSeries values = series[0];
        double total = 0D;
        for (int i = 0; i < categories.Count; i++) {
            if (TryGetSeriesValue(values, i, out double value) && value > 0D) {
                total += value;
            }
        }

        if (total <= 0D) {
            return;
        }

        double topCategoryLegendHeight = layout.LegendPosition == OfficeChartLegendPosition.Top
            ? GetCategoryLegendBandHeight(categories, width - 16D, layout)
            : 0D;
        IReadOnlyList<OfficeColor?> categoryPointColors = GetCategoryPointColors(style, values, categories.Count);
        if (topCategoryLegendHeight > 0D) {
            AddCategoryLegendBand(drawing, categories, 8D, contentTop + 2D, Math.Max(1D, width - 16D), style, layout, categoryPointColors, values.PointStyles);
            contentTop += topCategoryLegendHeight;
        }

        double categoryBottomLegendHeight = layout.LegendPosition == OfficeChartLegendPosition.Bottom
            ? GetCategoryLegendBandHeight(categories, width - 16D, layout)
            : bottomLegendHeight;
        double legendWidth = GetCategoryLegendWidth(categories, width, layout);
        bool leftLegend = layout.LegendPosition == OfficeChartLegendPosition.Left;
        GetRadialPlotGeometry(width, height, contentTop, categoryBottomLegendHeight, legendWidth,
            leftLegend, 48D, 36D, out double centerX, out double centerY, out double radius,
            out double visualWidth, out double contentHeight);
        radius /= 1D + GetMaximumRenderableExplosion(values, categories.Count) / 100D;
        double start = GetFirstSliceAngle(snapshot.RadialLayout);
        int zeroLabelIndex = 0;
        OfficeColor zeroLabelColor = GetPointDataLabelColor(style, values,
            Enumerable.Range(0, categories.Count).First(index => TryGetSeriesValue(values, index, out double firstValue) && firstValue > 0));
        for (int i = 0; i < categories.Count; i++) {
            if (!TryGetSeriesValue(values, i, out double seriesValue)) {
                continue;
            }

            double value = Math.Max(0D, seriesValue);
            double sweep = value / total * Math.PI * 2D;
            if (value > 0D) {
                double end = start + sweep;
                double middle = start + sweep / 2D;
                int explosion = values.PointExplosions?[i] ?? 0;
                double sliceCenterX = centerX + Math.Cos(middle) * radius * explosion / 100D;
                double sliceCenterY = centerY + Math.Sin(middle) * radius * explosion / 100D;
                var points = new List<OfficePoint> {
                    new OfficePoint(sliceCenterX, sliceCenterY)
                };
                int segments = Math.Max(2, (int)Math.Ceiling(sweep / (Math.PI / 18D)));
                for (int segment = 0; segment <= segments; segment++) {
                    double angle = start + sweep * segment / segments;
                    points.Add(new OfficePoint(
                        sliceCenterX + Math.Cos(angle) * radius,
                        sliceCenterY + Math.Sin(angle) * radius));
                }

                OfficeColor sliceColor = GetPointColor(style, values, i);
                AddStyledPointPolygon(drawing, points, sliceColor, GetPointStyle(values, i), OfficeColor.White, 0.5D);
                if (ShouldShowDataLabel(layout, 0, i)) {
                    AddPieDataLabel(drawing, layout, style, GetPointDataLabelColor(style, values, i), categories[i], values, value, total, sliceCenterX, sliceCenterY, radius * 0.58D, middle, zeroLabelIndex: null);
                }

                start = end;
            } else if (ShouldShowDataLabel(layout, 0, i)) {
                AddPieDataLabel(drawing, layout, style, zeroLabelColor, categories[i], values, 0D, total, centerX, centerY, radius * 0.9D, GetFirstSliceAngle(snapshot.RadialLayout), zeroLabelIndex);
                zeroLabelIndex++;
            }
        }

        if (layout.OverlayLegend) {
            AddOverlayCategoryLegend(drawing, categories, leftLegend ? legendWidth : 0D, contentTop + 4D, visualWidth, Math.Max(20D, contentHeight - 8D), style, layout, categoryPointColors, values.PointStyles);
        } else {
            AddCategoryLegend(
                drawing,
                categories,
                leftLegend ? 6D : width - legendWidth + 6D,
                contentTop + 12D,
                Math.Max(0D, legendWidth - 12D),
                Math.Max(20D, contentHeight - 24D),
                style,
                layout,
                categoryPointColors, values.PointStyles);
        }

        if (!layout.OverlayLegend && categoryBottomLegendHeight > 0D) {
            AddCategoryLegendBand(drawing, categories, 8D, height - categoryBottomLegendHeight + 2D, Math.Max(1D, width - 16D), style, layout, categoryPointColors, values.PointStyles);
        }
    }

    private static void AddDoughnutSeries(OfficeDrawing drawing, OfficeChartSnapshot snapshot, double width, double height, double contentTop, double bottomLegendHeight, OfficeChartStyle style, OfficeChartLayout layout) {
        IReadOnlyList<string> categories = snapshot.Data.Categories;
        IReadOnlyList<OfficeChartSeries> series = snapshot.Data.Series;
        var renderableSeries = new List<(OfficeChartSeries Series, int SourceIndex)>();
        for (int s = 0; s < series.Count; s++) {
            if (GetPositiveSeriesTotal(series[s], categories.Count) > 0D) {
                renderableSeries.Add((series[s], s));
            }
        }

        if (renderableSeries.Count == 0) {
            return;
        }

        double topCategoryLegendHeight = layout.LegendPosition == OfficeChartLegendPosition.Top
            ? GetCategoryLegendBandHeight(categories, width - 16D, layout)
            : 0D;
        OfficeChartSeries? legendSeries = GetCategoryLegendSeries(renderableSeries.ConvertAll(item => item.Series));
        IReadOnlyList<OfficeColor?>? legendPointColors = legendSeries == null ? null : GetCategoryPointColors(style, legendSeries, categories.Count);
        IReadOnlyList<OfficeChartPointStyle?>? legendPointStyles = legendSeries?.PointStyles;
        if (topCategoryLegendHeight > 0D) {
            AddCategoryLegendBand(drawing, categories, 8D, contentTop + 2D, Math.Max(1D, width - 16D), style, layout, legendPointColors, legendPointStyles);
            contentTop += topCategoryLegendHeight;
        }

        double categoryBottomLegendHeight = layout.LegendPosition == OfficeChartLegendPosition.Bottom
            ? GetCategoryLegendBandHeight(categories, width - 16D, layout)
            : bottomLegendHeight;
        double legendWidth = GetCategoryLegendWidth(categories, width, layout);
        bool leftLegend = layout.LegendPosition == OfficeChartLegendPosition.Left;
        GetRadialPlotGeometry(width, height, contentTop, categoryBottomLegendHeight, legendWidth,
            leftLegend, 48D, 36D, out double centerX, out double centerY, out double radius,
            out double visualWidth, out double contentHeight);

        int maximumExplosion = renderableSeries.Max(item =>
            GetMaximumRenderableExplosion(item.Series, categories.Count));
        radius /= 1D + maximumExplosion / 100D;
        double holeRadius = radius * snapshot.RadialLayout.DoughnutHolePercent / 100D;
        double ringThickness = (radius - holeRadius) / renderableSeries.Count;
        for (int s = 0; s < renderableSeries.Count; s++) {
            OfficeChartSeries values = renderableSeries[s].Series;
            int sourceSeriesIndex = renderableSeries[s].SourceIndex;
            double outerRadius = holeRadius + (s + 1) * ringThickness;
            double innerRadius = holeRadius + s * ringThickness;
            double total = GetPositiveSeriesTotal(values, categories.Count);
            double start = GetFirstSliceAngle(snapshot.RadialLayout);
            int zeroLabelIndex = 0;
            OfficeColor zeroLabelColor = GetPointDataLabelColor(style, values,
                Enumerable.Range(0, categories.Count).First(index => TryGetSeriesValue(values, index, out double firstValue) && firstValue > 0));
            for (int i = 0; i < categories.Count; i++) {
                if (!TryGetSeriesValue(values, i, out double seriesValue)) {
                    continue;
                }

                double value = Math.Max(0D, seriesValue);
                double sweep = value / total * Math.PI * 2D;
                if (value > 0D) {
                    double end = start + sweep;
                    double middle = start + sweep / 2D;
                    int explosion = values.PointExplosions?[i] ?? 0;
                    double sliceCenterX = centerX + Math.Cos(middle) * radius * explosion / 100D;
                    double sliceCenterY = centerY + Math.Sin(middle) * radius * explosion / 100D;
                    OfficeColor sliceColor = GetPointColor(style, values, i);
                    AddDoughnutSlice(drawing, sliceCenterX, sliceCenterY, outerRadius, innerRadius, start, sweep, sliceColor, GetPointStyle(values, i));
                    if (ShouldShowDataLabel(layout, sourceSeriesIndex, i)) {
                        AddPieDataLabel(drawing, layout, style, GetPointDataLabelColor(style, values, i), categories[i], values, value, total, sliceCenterX, sliceCenterY, (innerRadius + outerRadius) / 2D, middle, zeroLabelIndex: null);
                    }

                    start = end;
                } else if (s == 0 && ShouldShowDataLabel(layout, sourceSeriesIndex, i)) {
                    AddPieDataLabel(drawing, layout, style, zeroLabelColor, categories[i], values, 0D, total, centerX, centerY, (innerRadius + outerRadius) / 2D, GetFirstSliceAngle(snapshot.RadialLayout), zeroLabelIndex);
                    zeroLabelIndex++;
                }
            }
        }

        if (layout.OverlayLegend) {
            AddOverlayCategoryLegend(drawing, categories, leftLegend ? legendWidth : 0D, contentTop + 4D, visualWidth, Math.Max(20D, contentHeight - 8D), style, layout, legendPointColors, legendPointStyles);
        } else {
            AddCategoryLegend(
                drawing,
                categories,
                leftLegend ? 6D : width - legendWidth + 6D,
                contentTop + 12D,
                Math.Max(0D, legendWidth - 12D),
                Math.Max(20D, contentHeight - 24D),
                style,
                layout,
                legendPointColors, legendPointStyles);
        }

        if (!layout.OverlayLegend && categoryBottomLegendHeight > 0D) {
            AddCategoryLegendBand(drawing, categories, 8D, height - categoryBottomLegendHeight + 2D, Math.Max(1D, width - 16D), style, layout, legendPointColors, legendPointStyles);
        }
    }

    private static void AddPieDataLabel(
        OfficeDrawing drawing,
        OfficeChartLayout layout,
        OfficeChartStyle style,
        OfficeColor labelColor,
        string category,
        OfficeChartSeries series,
        double value,
        double total,
        double centerX,
        double centerY,
        double radius,
        double angle,
        int? zeroLabelIndex) {
        string label = FormatDataLabel(layout, category, series, value, total);
        if (string.IsNullOrWhiteSpace(label)) {
            return;
        }

        double labelWidth = Math.Min(78D, Math.Max(40D, label.Length * layout.DataLabelFontSize * 0.52D + 12D));
        double labelHeight = Math.Max(12D, layout.DataLabelFontSize + 6D);
        double distance = radius;
        double x = centerX + Math.Cos(angle) * distance - labelWidth / 2D;
        double y = centerY + Math.Sin(angle) * distance - labelHeight / 2D;
        if (zeroLabelIndex.HasValue) {
            y += zeroLabelIndex.Value * (labelHeight + 1D);
        }

        AddDataLabel(drawing, layout, style, label, x, y, labelWidth, labelHeight, OfficeTextAlignment.Center, labelColor);
    }

    private static double GetFirstSliceAngle(OfficeChartRadialLayout layout) =>
        (layout.FirstSliceAngleDegrees - 90D) * Math.PI / 180D;

    private static OfficeColor GetReadableDataLabelColor(OfficeColor fillColor) {
        double srgbR = fillColor.R / 255D;
        double srgbG = fillColor.G / 255D;
        double srgbB = fillColor.B / 255D;
        double luminance = 0.2126D * srgbR + 0.7152D * srgbG + 0.0722D * srgbB;
        return luminance < 0.52D ? OfficeColor.White : OfficeColor.Black;
    }

    private static double GetPositiveSeriesTotal(OfficeChartSeries values, int categoryCount) {
        double total = 0D;
        for (int i = 0; i < categoryCount; i++) {
            if (TryGetSeriesValue(values, i, out double value) && value > 0D) {
                total += value;
            }
        }

        return total;
    }

    private static int GetMaximumRenderableExplosion(OfficeChartSeries series, int categoryCount) {
        if (series.PointExplosions == null) return 0;
        int maximum = 0;
        for (int index = 0; index < categoryCount; index++)
            if (TryGetSeriesValue(series, index, out double value) && value > 0D)
                maximum = Math.Max(maximum, series.PointExplosions[index]);
        return maximum;
    }

    private static void AddPieSlice(OfficeDrawing drawing, double centerX, double centerY, double radius, double start, double sweep, OfficeColor color, OfficeChartPointStyle? pointStyle = null) {
        var points = new List<OfficePoint> {
            new OfficePoint(centerX, centerY)
        };
        int segments = Math.Max(2, (int)Math.Ceiling(sweep / (Math.PI / 18D)));
        for (int segment = 0; segment <= segments; segment++) {
            double angle = start + sweep * segment / segments;
            points.Add(new OfficePoint(
                centerX + Math.Cos(angle) * radius,
                centerY + Math.Sin(angle) * radius));
        }

        AddStyledPointPolygon(drawing, points, color, pointStyle, OfficeColor.White, 0.5D);
    }

    private static void AddDoughnutSlice(OfficeDrawing drawing, double centerX, double centerY, double outerRadius, double innerRadius, double start, double sweep, OfficeColor color, OfficeChartPointStyle? pointStyle = null) {
        if (innerRadius <= 0D) {
            AddPieSlice(drawing, centerX, centerY, outerRadius, start, sweep, color, pointStyle);
            return;
        }

        int segments = Math.Max(2, (int)Math.Ceiling(sweep / (Math.PI / 18D)));
        var points = new List<OfficePoint>((segments + 1) * 2);
        for (int segment = 0; segment <= segments; segment++) {
            double angle = start + sweep * segment / segments;
            points.Add(new OfficePoint(
                centerX + Math.Cos(angle) * outerRadius,
                centerY + Math.Sin(angle) * outerRadius));
        }

        for (int segment = segments; segment >= 0; segment--) {
            double angle = start + sweep * segment / segments;
            points.Add(new OfficePoint(
                centerX + Math.Cos(angle) * innerRadius,
                centerY + Math.Sin(angle) * innerRadius));
        }

        AddStyledPointPolygon(drawing, points, color, pointStyle, OfficeColor.White, 0.5D);
    }

}
