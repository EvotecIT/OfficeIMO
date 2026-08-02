using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointChart {
        /// <summary>
        /// Tries to create a dependency-free snapshot for rendering/export consumers.
        /// </summary>
        internal bool TryGetSnapshot(out PowerPointChartSnapshot snapshot) =>
            TryGetSnapshotWithOwnerColorScheme(
                forDataUpdate: false,
                out snapshot);

        private bool TryGetSnapshotForUpdate(out PowerPointChartSnapshot snapshot) =>
            TryGetSnapshotWithOwnerColorScheme(
                forDataUpdate: true,
                out snapshot);

        private bool TryGetSnapshotWithOwnerColorScheme(
            bool forDataUpdate,
            out PowerPointChartSnapshot snapshot) {
            try {
                return TryGetSnapshot(
                    GetOwnerColorScheme(),
                    forDataUpdate,
                    out snapshot);
            } catch {
                snapshot = null!;
                return false;
            }
        }

        private A.ColorScheme? GetOwnerColorScheme() {
            if (_ownerPart is SlidePart slidePart) {
                return slidePart.ThemeOverridePart?.ThemeOverride?.ColorScheme
                    ?? slidePart.SlideLayoutPart?.ThemeOverridePart?.ThemeOverride?
                        .ColorScheme
                    ?? slidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme?
                        .ThemeElements?.ColorScheme;
            }

            if (_ownerPart is SlideLayoutPart layoutPart) {
                return layoutPart.ThemeOverridePart?.ThemeOverride?.ColorScheme
                    ?? layoutPart.SlideMasterPart?.ThemePart?.Theme?.ThemeElements?
                        .ColorScheme;
            }

            if (_ownerPart is SlideMasterPart masterPart) {
                return masterPart.ThemePart?.Theme?.ThemeElements?.ColorScheme;
            }

            if (_ownerPart is NotesSlidePart notesPart) {
                return notesPart.ThemeOverridePart?.ThemeOverride?.ColorScheme
                    ?? notesPart.NotesMasterPart?.ThemePart?.Theme?.ThemeElements?
                        .ColorScheme;
            }

            if (_ownerPart is NotesMasterPart notesMasterPart) {
                return notesMasterPart.ThemePart?.Theme?.ThemeElements?.ColorScheme;
            }

            return (_ownerPart as HandoutMasterPart)?.ThemePart?.Theme?
                .ThemeElements?.ColorScheme;
        }

        internal bool TryGetSnapshot(A.ColorScheme? colorScheme,
            out PowerPointChartSnapshot snapshot) =>
            TryGetSnapshot(colorScheme, forDataUpdate: false, out snapshot);

        private bool TryGetSnapshot(A.ColorScheme? colorScheme,
            bool forDataUpdate, out PowerPointChartSnapshot snapshot) {
            try {
                ChartPart chartPart = GetChartPart();
                C.Chart? chart = chartPart.ChartSpace?.GetFirstChild<C.Chart>();
                C.PlotArea? plotArea = chart?.GetFirstChild<C.PlotArea>();
                if (chart == null || plotArea == null) {
                    snapshot = null!;
                    return false;
                }

                if (!forDataUpdate
                    && !TryReadSharedTextStyle(chart, out _)) {
                    snapshot = null!;
                    return false;
                }

                if (HasUnsupportedChartGroupElements(plotArea)) {
                    snapshot = null!;
                    return false;
                }

                if (TryCreateMixedChartSnapshot(chart, plotArea, colorScheme, out snapshot)) {
                    return true;
                }

                if (CountSupportedChartElements(plotArea) > 1) {
                    snapshot = null!;
                    return false;
                }

                if (plotArea.GetFirstChild<C.BarChart>() is C.BarChart barChart) {
                    PowerPointChartSnapshotKind kind = GetBarChartSnapshotKind(barChart);
                    PowerPointChartData? data = ReadCategorySeriesData(barChart.Elements<C.BarChartSeries>().Cast<OpenXmlCompositeElement>(), kind, colorScheme);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, kind, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.LineChart>() is C.LineChart lineChart) {
                    PowerPointChartSnapshotKind kind = GetLineChartSnapshotKind(lineChart);
                    PowerPointChartData? data = ReadCategorySeriesData(lineChart.Elements<C.LineChartSeries>().Cast<OpenXmlCompositeElement>(), kind, colorScheme);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, kind, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.AreaChart>() is C.AreaChart areaChart) {
                    PowerPointChartSnapshotKind kind = GetAreaChartSnapshotKind(areaChart);
                    PowerPointChartData? data = ReadCategorySeriesData(areaChart.Elements<C.AreaChartSeries>().Cast<OpenXmlCompositeElement>(), kind, colorScheme);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, kind, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.RadarChart>() is C.RadarChart radarChart) {
                    PowerPointChartData? data = ReadCategorySeriesData(radarChart.Elements<C.RadarChartSeries>().Cast<OpenXmlCompositeElement>(), PowerPointChartSnapshotKind.Radar, colorScheme);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, PowerPointChartSnapshotKind.Radar, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.ScatterChart>() is C.ScatterChart scatterChart) {
                    PowerPointChartData? data = ReadScatterSeriesData(scatterChart.Elements<C.ScatterChartSeries>(), colorScheme);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, PowerPointChartSnapshotKind.Scatter, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.BubbleChart>() is C.BubbleChart bubbleChart) {
                    if (!forDataUpdate &&
                        (IsVaryColorsEnabled(
                             bubbleChart.GetFirstChild<C.VaryColors>()) ||
                         IsBubble3DEnabled(
                             bubbleChart.GetFirstChild<C.Bubble3D>()) ||
                         HasUnsupportedBubbleSourceVisibility(chartPart, chart) ||
                         HasUnsupportedBubbleAxes(plotArea, bubbleChart) ||
                         HasUnsupportedBubbleLegend(chart) ||
                         HasUnsupportedBubbleAreaLayout(chartPart, plotArea) ||
                         HasEnabledBubbleDataLabels(bubbleChart) ||
                         bubbleChart.Elements<C.BubbleChartSeries>().Any(series =>
                             IsBubble3DEnabled(
                                 series.GetFirstChild<C.Bubble3D>()) ||
                             series.Elements<C.DataPoint>().Any(point =>
                                 IsBubble3DEnabled(
                                     point.GetFirstChild<C.Bubble3D>()))))) {
                        snapshot = null!;
                        return false;
                    }
                    PowerPointChartData? data = ReadBubbleSeriesData(
                        bubbleChart.Elements<C.BubbleChartSeries>(), colorScheme,
                        forDataUpdate);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    uint bubbleScale = bubbleChart.GetFirstChild<C.BubbleScale>()?.Val?.Value ?? 100U;
                    if (bubbleScale > 300U) {
                        snapshot = null!;
                        return false;
                    }
                    OfficeChartBubbleSizeMode bubbleSizeMode =
                        bubbleChart.GetFirstChild<C.SizeRepresents>()?.Val?.Value ==
                        C.SizeRepresentsValues.Width
                            ? OfficeChartBubbleSizeMode.Width
                            : OfficeChartBubbleSizeMode.Area;
                    snapshot = CreateSnapshot(chart, PowerPointChartSnapshotKind.Bubble, data,
                        bubbleSizeMode, bubbleScale);
                    return true;
                }

                if (plotArea.GetFirstChild<C.PieChart>() is C.PieChart pieChart) {
                    PowerPointChartData? data = ReadCategorySeriesData(pieChart.Elements<C.PieChartSeries>().Cast<OpenXmlCompositeElement>(), PowerPointChartSnapshotKind.Pie, colorScheme);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, PowerPointChartSnapshotKind.Pie, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.DoughnutChart>() is C.DoughnutChart doughnutChart) {
                    PowerPointChartData? data = ReadCategorySeriesData(doughnutChart.Elements<C.PieChartSeries>().Cast<OpenXmlCompositeElement>(), PowerPointChartSnapshotKind.Doughnut, colorScheme);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, PowerPointChartSnapshotKind.Doughnut, data);
                    return true;
                }

                snapshot = null!;
                return false;
            } catch {
                snapshot = null!;
                return false;
            }
        }

        private static bool IsBubble3DEnabled(C.Bubble3D? bubble3D) =>
            bubble3D != null && bubble3D.Val?.Value != false;

        private static bool IsVaryColorsEnabled(C.VaryColors? varyColors) =>
            varyColors != null && varyColors.Val?.Value != false;

        private static bool HasUnsupportedBubbleAxes(
            C.PlotArea plotArea, C.BubbleChart chart) {
            if (!TryGetReferencedBubbleAxes(
                    plotArea, chart, out C.ValueAxis horizontalAxis,
                    out C.ValueAxis verticalAxis)) {
                return true;
            }
            if (horizontalAxis.AxisPosition?.Val?.Value !=
                    C.AxisPositionValues.Bottom ||
                verticalAxis.AxisPosition?.Val?.Value !=
                    C.AxisPositionValues.Left) {
                return true;
            }
            if (!HasSupportedDefaultBubbleGridlines(
                    horizontalAxis, verticalAxis)) {
                return true;
            }
            return new[] { horizontalAxis, verticalAxis }.Any(axis =>
                HasUnsupportedBubbleAxisPresentation(axis) ||
                (axis.GetFirstChild<C.Delete>() is C.Delete delete &&
                  delete.Val?.Value != false) ||
                 axis.GetFirstChild<C.MajorUnit>() != null ||
                 axis.GetFirstChild<C.MinorUnit>() != null ||
                 axis.GetFirstChild<C.DisplayUnits>() != null ||
                 axis.GetFirstChild<C.CrossesAt>() != null ||
                 HasUnsupportedSharedAxisNumberFormat(axis) ||
                 (axis.GetFirstChild<C.TickLabelPosition>() is
                      C.TickLabelPosition tickLabelPosition &&
                  tickLabelPosition.Val?.Value !=
                      C.TickLabelPositionValues.NextTo) ||
                 (axis.GetFirstChild<C.Crosses>() is C.Crosses crosses &&
                  crosses.Val?.Value != C.CrossesValues.AutoZero) ||
                 (axis.GetFirstChild<C.Scaling>() is C.Scaling scaling &&
                  (scaling.GetFirstChild<C.LogBase>() != null ||
                   scaling.GetFirstChild<C.MinAxisValue>() != null ||
                   scaling.GetFirstChild<C.MaxAxisValue>() != null ||
                   scaling.GetFirstChild<C.Orientation>()?.Val?.Value ==
                      C.OrientationValues.MaxMin)));
        }

        private static bool HasUnsupportedBubbleAxisPresentation(
            C.ValueAxis axis) =>
            HasUnsupportedBubbleTitle(axis.GetFirstChild<C.Title>()) ||
            HasUnsupportedBubbleTextStyle(axis) ||
            HasUnsupportedBubbleShapeProperties(axis);

        private static bool HasSupportedDefaultBubbleGridlines(
            C.ValueAxis horizontalAxis, C.ValueAxis verticalAxis) {
            if (horizontalAxis.GetFirstChild<C.MajorGridlines>() != null ||
                horizontalAxis.GetFirstChild<C.MinorGridlines>() != null ||
                verticalAxis.GetFirstChild<C.MinorGridlines>() != null) {
                return false;
            }

            C.MajorGridlines? gridlines =
                verticalAxis.GetFirstChild<C.MajorGridlines>();
            C.ChartShapeProperties? properties =
                gridlines?.GetFirstChild<C.ChartShapeProperties>();
            A.Outline? outline = properties?.GetFirstChild<A.Outline>();
            if (gridlines == null || properties == null || outline == null ||
                gridlines.ChildElements.Count != 1 ||
                properties.ChildElements.Count != 1 ||
                outline.ChildElements.Count != 1 ||
                outline.Width?.Value !=
                    PowerPointUnits.FromPoints(0.5D)) {
                return false;
            }

            OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(
                outline.GetFirstChild<A.SolidFill>(), colorScheme: null);
            return color == OfficeChartStyle.Default.GridLineColor;
        }

        private static bool TryGetReferencedBubbleAxes(
            C.PlotArea plotArea, C.BubbleChart chart,
            out C.ValueAxis horizontalAxis, out C.ValueAxis verticalAxis) {
            horizontalAxis = null!;
            verticalAxis = null!;
            List<C.AxisId> references =
                chart.Elements<C.AxisId>().ToList();
            if (references.Count != 2 ||
                references.Any(axis => axis.Val?.Value == null)) {
                return false;
            }
            uint horizontalId = references[0].Val!.Value;
            uint verticalId = references[1].Val!.Value;
            if (horizontalId == verticalId) return false;
            C.ValueAxis? horizontal = plotArea.Elements<C.ValueAxis>()
                .FirstOrDefault(axis =>
                    axis.AxisId?.Val?.Value == horizontalId);
            C.ValueAxis? vertical = plotArea.Elements<C.ValueAxis>()
                .FirstOrDefault(axis =>
                    axis.AxisId?.Val?.Value == verticalId);
            if (horizontal == null || vertical == null) return false;
            horizontalAxis = horizontal;
            verticalAxis = vertical;
            return true;
        }

        private static bool HasUnsupportedBubbleLegend(C.Chart chart) {
            C.Legend? legend = chart.GetFirstChild<C.Legend>();
            return legend != null &&
                (legend.GetFirstChild<C.LegendPosition>()?.Val?.Value ==
                     C.LegendPositionValues.TopRight ||
                 legend.GetFirstChild<C.Layout>()?
                     .GetFirstChild<C.ManualLayout>() != null ||
                 HasUnsupportedBubbleTextStyle(legend) ||
                 HasUnsupportedBubbleShapeProperties(legend));
        }

        private static bool HasUnsupportedBubbleAreaLayout(
            ChartPart chartPart, C.PlotArea plotArea) =>
            HasUnsupportedBubbleTitle(
                chartPart.ChartSpace?.GetFirstChild<C.Chart>()?
                    .GetFirstChild<C.Title>()) ||
            plotArea.GetFirstChild<C.Layout>()?
                .GetFirstChild<C.ManualLayout>() != null ||
            chartPart.ChartSpace?.GetFirstChild<C.ShapeProperties>()?
                .ChildElements.Count > 0 ||
            plotArea.GetFirstChild<C.ShapeProperties>()?
                .ChildElements.Count > 0;

        private static bool HasUnsupportedBubbleTitle(C.Title? title) =>
            title != null &&
            (title.GetFirstChild<C.Layout>()?
                 .GetFirstChild<C.ManualLayout>() != null ||
             HasUnsupportedBubbleTextStyle(title) ||
             HasUnsupportedBubbleShapeProperties(title));

        private static bool HasUnsupportedBubbleTextStyle(
            OpenXmlElement parent) =>
            parent.Descendants<A.RunProperties>()
                .Any(HasUnsupportedBubbleTextCharacterProperties) ||
            parent.Descendants<A.DefaultRunProperties>()
                .Any(HasUnsupportedBubbleTextCharacterProperties) ||
            parent.Descendants<A.EndParagraphRunProperties>()
                .Any(HasUnsupportedBubbleTextCharacterProperties) ||
            parent.Descendants<A.BodyProperties>()
                .Any(properties =>
                    properties.HasAttributes ||
                    properties.ChildElements.Count > 0) ||
            parent.Descendants<A.ListStyle>()
                .Any(style => style.ChildElements.Count > 0) ||
            parent.Descendants<A.ParagraphProperties>()
                .Any(properties =>
                    properties.HasAttributes ||
                    properties.ChildElements.Any(child =>
                        child is not A.DefaultRunProperties));

        private static bool HasUnsupportedBubbleTextCharacterProperties(
            A.TextCharacterPropertiesType properties) =>
            properties.ChildElements.Count > 0 ||
            properties.GetAttributes().Any(attribute =>
                !string.Equals(
                    attribute.LocalName, "lang",
                    StringComparison.Ordinal));

        private static bool HasUnsupportedBubbleShapeProperties(
            OpenXmlElement parent) {
            C.ChartShapeProperties? properties =
                parent.GetFirstChild<C.ChartShapeProperties>();
            return properties != null &&
                (properties.HasAttributes ||
                 properties.ChildElements.Count > 0);
        }

        private static bool HasEnabledBubbleDataLabels(C.BubbleChart chart) =>
            chart.Descendants<C.ShowLegendKey>().Any(item => item.Val?.Value != false) ||
            chart.Descendants<C.ShowValue>().Any(item => item.Val?.Value != false) ||
            chart.Descendants<C.ShowCategoryName>().Any(item => item.Val?.Value != false) ||
            chart.Descendants<C.ShowSeriesName>().Any(item => item.Val?.Value != false) ||
            chart.Descendants<C.ShowPercent>().Any(item => item.Val?.Value != false) ||
            chart.Descendants<C.ShowBubbleSize>().Any(item => item.Val?.Value != false) ||
            chart.Descendants<C.DataLabel>().Any(label => {
                C.Delete? delete = label.GetFirstChild<C.Delete>();
                return label.GetFirstChild<C.ChartText>() != null &&
                    (delete == null || delete.Val?.Value == false);
            });

        private static int CountSupportedChartElements(C.PlotArea plotArea) {
            return plotArea.Elements<C.BarChart>().Count()
                + plotArea.Elements<C.LineChart>().Count()
                + plotArea.Elements<C.AreaChart>().Count()
                + plotArea.Elements<C.RadarChart>().Count()
                + plotArea.Elements<C.ScatterChart>().Count()
                + plotArea.Elements<C.BubbleChart>().Count()
                + plotArea.Elements<C.PieChart>().Count()
                + plotArea.Elements<C.DoughnutChart>().Count();
        }

        private static bool HasUnsupportedChartGroupElements(C.PlotArea plotArea) =>
            plotArea.ChildElements.Any(element =>
                element.LocalName.EndsWith("Chart", StringComparison.Ordinal) &&
                element is not C.BarChart &&
                element is not C.LineChart &&
                element is not C.AreaChart &&
                element is not C.RadarChart &&
                element is not C.ScatterChart &&
                element is not C.BubbleChart &&
                element is not C.PieChart &&
                element is not C.DoughnutChart);

        private bool TryCreateMixedChartSnapshot(C.Chart chart, C.PlotArea plotArea, A.ColorScheme? colorScheme, out PowerPointChartSnapshot snapshot) {
            snapshot = null!;
            if (CountSupportedChartElements(plotArea) <= 1) {
                return false;
            }
            if (plotArea.Elements<C.BubbleChart>().Any()) {
                return false;
            }

            var parts = new List<(PowerPointChartSnapshotKind Kind, PowerPointChartData Data)>();
            foreach (OpenXmlElement element in plotArea.ChildElements) {
                if (element is C.BarChart barChart) {
                    PowerPointChartSnapshotKind kind = GetBarChartSnapshotKind(barChart);
                    PowerPointChartData? data = ReadCategorySeriesData(
                        barChart.Elements<C.BarChartSeries>().Cast<OpenXmlCompositeElement>(), kind, colorScheme,
                        GetAxisGroup(plotArea, barChart));
                    if (data != null) {
                        parts.Add((kind, data));
                    }
                } else if (element is C.LineChart lineChart) {
                    PowerPointChartSnapshotKind kind = GetLineChartSnapshotKind(lineChart);
                    PowerPointChartData? data = ReadCategorySeriesData(
                        lineChart.Elements<C.LineChartSeries>().Cast<OpenXmlCompositeElement>(), kind, colorScheme,
                        GetAxisGroup(plotArea, lineChart));
                    if (data != null) {
                        parts.Add((kind, data));
                    }
                } else if (element is C.AreaChart areaChart) {
                    PowerPointChartSnapshotKind kind = GetAreaChartSnapshotKind(areaChart);
                    PowerPointChartData? data = ReadCategorySeriesData(
                        areaChart.Elements<C.AreaChartSeries>().Cast<OpenXmlCompositeElement>(), kind, colorScheme,
                        GetAxisGroup(plotArea, areaChart));
                    if (data != null) {
                        parts.Add((kind, data));
                    }
                } else if (element is C.ScatterChart scatterChart) {
                    PowerPointChartData? data = ReadScatterSeriesData(scatterChart.Elements<C.ScatterChartSeries>(), colorScheme);
                    if (data != null) {
                        parts.Add((PowerPointChartSnapshotKind.Scatter, data));
                    }
                }
            }

            if (parts.Count <= 1) {
                return false;
            }

            if (parts.Any(part => part.Kind == PowerPointChartSnapshotKind.Scatter) &&
                parts.Any(part => part.Kind != PowerPointChartSnapshotKind.Scatter)) {
                return false;
            }

            if (parts.Any(part => IsHorizontalBarKind(part.Kind)) &&
                parts.Any(part => !IsHorizontalBarKind(part.Kind))) {
                return false;
            }

            IReadOnlyList<string> categories = parts[0].Data.Categories;
            var series = new List<PowerPointChartSeries>();
            foreach (var part in parts) {
                foreach (PowerPointChartSeries item in part.Data.Series) {
                    if (item.Values.Count == categories.Count || HasAlignedScatterPoints(item)) {
                        series.Add(item);
                    }
                }
            }

            if (series.Count == 0) {
                return false;
            }

            snapshot = CreateSnapshot(chart, parts[0].Kind, new PowerPointChartData(categories, series));
            return true;
        }

        private static bool HasAlignedScatterPoints(PowerPointChartSeries series) =>
            series.XValues != null &&
            series.XValues.Count == series.Values.Count &&
            series.Values.Count > 0;

        private static bool IsHorizontalBarKind(PowerPointChartSnapshotKind kind) =>
            kind == PowerPointChartSnapshotKind.ClusteredBar ||
            kind == PowerPointChartSnapshotKind.StackedBar ||
            kind == PowerPointChartSnapshotKind.StackedBar100;

        private static OfficeChartAxisGroup GetAxisGroup(C.PlotArea plotArea, OpenXmlCompositeElement chart) {
            HashSet<uint> axisIds = new(chart.Elements<C.AxisId>()
                .Where(axis => axis.Val?.Value != null).Select(axis => axis.Val!.Value));
            return plotArea.Elements<C.ValueAxis>().Any(axis =>
                       axis.AxisId?.Val?.Value != null && axisIds.Contains(axis.AxisId.Val.Value) &&
                       (axis.AxisPosition?.Val?.Value == C.AxisPositionValues.Right ||
                        axis.AxisPosition?.Val?.Value == C.AxisPositionValues.Top))
                ? OfficeChartAxisGroup.Secondary
                : OfficeChartAxisGroup.Primary;
        }

        private PowerPointChartSnapshot CreateSnapshot(C.Chart chart,
            PowerPointChartSnapshotKind kind, PowerPointChartData data,
            OfficeChartBubbleSizeMode bubbleSizeMode = OfficeChartBubbleSizeMode.Area,
            double bubbleScalePercent = 100D) {
            HashSet<uint> hiddenLegendSeries = GetHiddenLegendSeriesIndexes(chart);
            bool hasLegend = chart.GetFirstChild<C.Legend>() != null;
            for (int seriesIndex = 0; seriesIndex < data.Series.Count; seriesIndex++) {
                PowerPointChartSeries series = data.Series[seriesIndex];
                uint sourceIndex = series.SourceIndex ?? (uint)seriesIndex;
                uint legendIndex = kind == PowerPointChartSnapshotKind.Bubble
                    ? (uint)seriesIndex
                    : sourceIndex;
                series.ShowInLegend = hasLegend &&
                    !hiddenLegendSeries.Contains(legendIndex);
            }

            return new PowerPointChartSnapshot(
                Name ?? string.Empty,
                ReadTitle(chart),
                kind,
                data,
                WidthPoints,
                HeightPoints,
                bubbleSizeMode,
                bubbleScalePercent,
                ReadChartLayout(chart, kind),
                ReadSharedTextStyle(chart));
        }

        private static OfficeChartLayout ReadChartLayout(
            C.Chart chart, PowerPointChartSnapshotKind kind) {
            C.Legend? legend = chart.GetFirstChild<C.Legend>();
            C.LegendPositionValues? nativePosition =
                legend?.GetFirstChild<C.LegendPosition>()?.Val?.Value;
            OfficeChartLegendPosition position =
                nativePosition == C.LegendPositionValues.Left
                    ? OfficeChartLegendPosition.Left
                    : nativePosition == C.LegendPositionValues.Top
                        ? OfficeChartLegendPosition.Top
                        : nativePosition == C.LegendPositionValues.Bottom
                            ? OfficeChartLegendPosition.Bottom
                            : OfficeChartLegendPosition.Right;
            bool overlay = legend?.GetFirstChild<C.Overlay>() is C.Overlay item &&
                item.Val?.Value != false;
            bool overlayTitle =
                chart.GetFirstChild<C.Title>()?.GetFirstChild<C.Overlay>()
                    is C.Overlay titleOverlay &&
                titleOverlay.Val?.Value != false;

            string? horizontalAxisTitle = null;
            string? verticalAxisTitle = null;
            string? horizontalAxisNumberFormat = null;
            string? verticalAxisNumberFormat = null;
            OfficeChartAxisTickMark horizontalMajorTickMark =
                OfficeChartAxisTickMark.None;
            OfficeChartAxisTickMark verticalMajorTickMark =
                OfficeChartAxisTickMark.None;
            OfficeChartAxisTickMark horizontalMinorTickMark =
                OfficeChartAxisTickMark.None;
            OfficeChartAxisTickMark verticalMinorTickMark =
                OfficeChartAxisTickMark.None;
            if (kind == PowerPointChartSnapshotKind.Bubble &&
                chart.GetFirstChild<C.PlotArea>() is C.PlotArea plotArea &&
                plotArea.GetFirstChild<C.BubbleChart>() is C.BubbleChart bubble &&
                TryGetReferencedBubbleAxes(
                    plotArea, bubble, out C.ValueAxis horizontalAxis,
                    out C.ValueAxis verticalAxis)) {
                horizontalAxisTitle = ReadAxisTitle(horizontalAxis);
                verticalAxisTitle = ReadAxisTitle(verticalAxis);
                horizontalAxisNumberFormat =
                    ReadAxisNumberFormat(horizontalAxis);
                verticalAxisNumberFormat =
                    ReadAxisNumberFormat(verticalAxis);
                horizontalMajorTickMark = ReadAxisTickMark(
                    horizontalAxis.GetFirstChild<C.MajorTickMark>()?
                        .Val?.Value);
                verticalMajorTickMark = ReadAxisTickMark(
                    verticalAxis.GetFirstChild<C.MajorTickMark>()?
                        .Val?.Value);
                horizontalMinorTickMark = ReadAxisTickMark(
                    horizontalAxis.GetFirstChild<C.MinorTickMark>()?
                        .Val?.Value);
                verticalMinorTickMark = ReadAxisTickMark(
                    verticalAxis.GetFirstChild<C.MinorTickMark>()?
                        .Val?.Value);
            }

            return new OfficeChartLayout(overlayLegend: overlay,
                overlayTitle: overlayTitle,
                showLegend: legend != null,
                legendPosition: position,
                categoryAxisTitle: horizontalAxisTitle,
                valueAxisTitle: verticalAxisTitle,
                horizontalAxisNumberFormat: horizontalAxisNumberFormat,
                verticalAxisNumberFormat: verticalAxisNumberFormat,
                horizontalAxisMajorTickMark: horizontalMajorTickMark,
                verticalAxisMajorTickMark: verticalMajorTickMark,
                horizontalAxisMinorTickMark: horizontalMinorTickMark,
                verticalAxisMinorTickMark: verticalMinorTickMark);
        }

        private static OfficeChartAxisTickMark ReadAxisTickMark(
            C.TickMarkValues? value) =>
            value == C.TickMarkValues.Inside
                ? OfficeChartAxisTickMark.Inside
                : value == C.TickMarkValues.Outside
                    ? OfficeChartAxisTickMark.Outside
                    : value == C.TickMarkValues.Cross
                        ? OfficeChartAxisTickMark.Cross
                        : OfficeChartAxisTickMark.None;

        private static string? ReadAxisTitle(C.ValueAxis axis) =>
            ReadChartText(
                axis.GetFirstChild<C.Title>()?.GetFirstChild<C.ChartText>());

        private static string? ReadAxisNumberFormat(C.ValueAxis axis) {
            string? format = axis.GetFirstChild<C.NumberingFormat>()?
                .FormatCode?.Value;
            return string.IsNullOrWhiteSpace(format) ? null : format;
        }

        private static bool HasUnsupportedSharedAxisNumberFormat(
            C.ValueAxis axis) {
            string? format = ReadAxisNumberFormat(axis);
            if (string.IsNullOrWhiteSpace(format)) return false;
            if (string.Equals(format, "General",
                    StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            bool inQuotedLiteral = false;
            bool escaped = false;
            bool sectionHasPlaceholder = false;
            for (int index = 0; index < format!.Length; index++) {
                char value = format[index];
                if (escaped) {
                    escaped = false;
                    continue;
                }
                if (value == '\\') {
                    escaped = true;
                    continue;
                }
                if (value == '"') {
                    inQuotedLiteral = !inQuotedLiteral;
                    continue;
                }
                if (inQuotedLiteral) {
                    continue;
                }
                if (value == '0' || value == '#' || value == '?') {
                    sectionHasPlaceholder = true;
                    continue;
                }
                if (value == ';') {
                    if (!sectionHasPlaceholder) return true;
                    sectionHasPlaceholder = false;
                    continue;
                }
                if (value == '/' || value == '@' ||
                    value == '[' || value == ']') {
                    return true;
                }
                if (value != 'E' && value != 'e') continue;

                int next = index + 1;
                if (next < format.Length &&
                    (format[next] == '+' || format[next] == '-')) {
                    next++;
                }
                if (next < format.Length &&
                    (format[next] == '0' || format[next] == '#' ||
                     format[next] == '?')) {
                    return true;
                }
            }

            return inQuotedLiteral || escaped || !sectionHasPlaceholder;
        }

        private static PowerPointChartSnapshotKind GetBarChartSnapshotKind(C.BarChart chart) {
            C.BarDirectionValues direction = chart.GetFirstChild<C.BarDirection>()?.Val?.Value ?? C.BarDirectionValues.Column;
            C.BarGroupingValues grouping = chart.GetFirstChild<C.BarGrouping>()?.Val?.Value ?? C.BarGroupingValues.Clustered;
            bool horizontal = direction == C.BarDirectionValues.Bar;

            if (grouping == C.BarGroupingValues.Stacked) {
                return horizontal ? PowerPointChartSnapshotKind.StackedBar : PowerPointChartSnapshotKind.StackedColumn;
            }

            if (grouping == C.BarGroupingValues.PercentStacked) {
                return horizontal ? PowerPointChartSnapshotKind.StackedBar100 : PowerPointChartSnapshotKind.StackedColumn100;
            }

            return horizontal ? PowerPointChartSnapshotKind.ClusteredBar : PowerPointChartSnapshotKind.ClusteredColumn;
        }

        private static PowerPointChartSnapshotKind GetLineChartSnapshotKind(C.LineChart chart) {
            C.GroupingValues grouping = chart.GetFirstChild<C.Grouping>()?.Val?.Value ?? C.GroupingValues.Standard;
            if (grouping == C.GroupingValues.Stacked) {
                return PowerPointChartSnapshotKind.StackedLine;
            }

            if (grouping == C.GroupingValues.PercentStacked) {
                return PowerPointChartSnapshotKind.StackedLine100;
            }

            return PowerPointChartSnapshotKind.Line;
        }

        private static PowerPointChartSnapshotKind GetAreaChartSnapshotKind(C.AreaChart chart) {
            C.GroupingValues grouping = chart.GetFirstChild<C.Grouping>()?.Val?.Value ?? C.GroupingValues.Standard;
            if (grouping == C.GroupingValues.Stacked) {
                return PowerPointChartSnapshotKind.StackedArea;
            }

            if (grouping == C.GroupingValues.PercentStacked) {
                return PowerPointChartSnapshotKind.StackedArea100;
            }

            return PowerPointChartSnapshotKind.Area;
        }

        private static PowerPointChartData? ReadCategorySeriesData(IEnumerable<OpenXmlCompositeElement> seriesElements,
            PowerPointChartSnapshotKind? chartKind = null, A.ColorScheme? colorScheme = null,
            OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary) {
            var seriesList = seriesElements.ToList();
            if (seriesList.Count == 0) {
                return null;
            }

            IReadOnlyList<string> categories = Array.Empty<string>();
            for (int i = 0; i < seriesList.Count; i++) {
                IReadOnlyList<double> values = ReadCachedNumbers(seriesList[i].GetFirstChild<C.Values>());
                if (values.Count == 0) {
                    continue;
                }

                categories = ReadCachedStrings(seriesList[i].GetFirstChild<C.CategoryAxisData>());
                if (categories.Count == 0) {
                    categories = CreateFallbackCategories(values.Count);
                }

                if (categories.Count > 0) {
                    break;
                }
            }

            if (categories.Count == 0) {
                return null;
            }

            var series = new List<PowerPointChartSeries>();
            for (int i = 0; i < seriesList.Count; i++) {
                OpenXmlCompositeElement seriesElement = seriesList[i];
                IReadOnlyList<double> values = NormalizeValues(ReadCachedNumbers(seriesElement.GetFirstChild<C.Values>()), categories.Count);
                if (values.Count == 0) {
                    continue;
                }

                string name = ReadSeriesName(seriesElement);
                if (string.IsNullOrWhiteSpace(name)) {
                    name = "Series " + (i + 1).ToString(CultureInfo.InvariantCulture);
                }

                series.Add(new PowerPointChartSeries(name, values, null, chartKind,
                    ReadSeriesColor(seriesElement, chartKind, colorScheme), ReadSeriesStrokeWidth(seriesElement),
                    axisGroup) {
                    SourceIndex = seriesElement.GetFirstChild<C.Index>()?.Val?.Value
                });
            }

            return series.Count == 0 ? null : new PowerPointChartData(categories, series);
        }

        private static PowerPointChartData? ReadScatterSeriesData(IEnumerable<C.ScatterChartSeries> seriesElements, A.ColorScheme? colorScheme = null) {
            var seriesList = seriesElements.ToList();
            if (seriesList.Count == 0) {
                return null;
            }

            var series = new List<PowerPointChartSeries>();
            IReadOnlyList<double>? categoryXValues = null;
            for (int i = 0; i < seriesList.Count; i++) {
                C.ScatterChartSeries seriesElement = seriesList[i];
                IReadOnlyList<double> xValues = ReadCachedNumbers(seriesElement.GetFirstChild<C.XValues>());
                IReadOnlyList<double> yValues = ReadCachedNumbers(seriesElement.GetFirstChild<C.YValues>());
                int pointCount = Math.Min(xValues.Count, yValues.Count);
                if (pointCount == 0) {
                    continue;
                }

                IReadOnlyList<double> values = NormalizeValues(yValues, pointCount);
                if (values.Count == 0) {
                    continue;
                }

                categoryXValues ??= xValues.Take(pointCount).ToList();
                string name = ReadSeriesName(seriesElement);
                if (string.IsNullOrWhiteSpace(name)) {
                    name = "Series " + (i + 1).ToString(CultureInfo.InvariantCulture);
                }

                series.Add(new PowerPointChartSeries(name, values, xValues.Take(pointCount).ToList(),
                    PowerPointChartSnapshotKind.Scatter,
                    ReadSeriesColor(seriesElement, PowerPointChartSnapshotKind.Scatter, colorScheme),
                    ReadSeriesStrokeWidth(seriesElement)) {
                    SourceIndex = seriesElement.GetFirstChild<C.Index>()?.Val?.Value
                });
            }

            if (series.Count == 0 || categoryXValues == null || categoryXValues.Count == 0) {
                return null;
            }

            var categories = categoryXValues
                .Select(value => value.ToString(CultureInfo.InvariantCulture))
                .ToList();
            return series.Count == 0 ? null : new PowerPointChartData(categories, series);
        }

        private static OfficeColor? ReadSeriesColor(OpenXmlCompositeElement seriesElement, PowerPointChartSnapshotKind? chartKind, A.ColorScheme? colorScheme) {
            C.ChartShapeProperties? properties = seriesElement.GetFirstChild<C.ChartShapeProperties>();
            if (properties == null) {
                return null;
            }

            OfficeColor? fillColor = OfficeOpenXmlThemeColorResolver.ResolveColor(properties.GetFirstChild<A.SolidFill>(), colorScheme);
            if (IsFilledChartKind(chartKind)) {
                return fillColor;
            }

            OfficeColor? lineColor = OfficeOpenXmlThemeColorResolver.ResolveColor(properties.GetFirstChild<A.Outline>()?.GetFirstChild<A.SolidFill>(), colorScheme);
            if (lineColor.HasValue) {
                return lineColor;
            }

            return fillColor;
        }

        private static bool IsFilledChartKind(PowerPointChartSnapshotKind? chartKind) =>
            chartKind == PowerPointChartSnapshotKind.ClusteredColumn ||
            chartKind == PowerPointChartSnapshotKind.StackedColumn ||
            chartKind == PowerPointChartSnapshotKind.StackedColumn100 ||
            chartKind == PowerPointChartSnapshotKind.ClusteredBar ||
            chartKind == PowerPointChartSnapshotKind.StackedBar ||
            chartKind == PowerPointChartSnapshotKind.StackedBar100 ||
            chartKind == PowerPointChartSnapshotKind.Area ||
            chartKind == PowerPointChartSnapshotKind.StackedArea ||
            chartKind == PowerPointChartSnapshotKind.StackedArea100 ||
            chartKind == PowerPointChartSnapshotKind.Bubble ||
            chartKind == PowerPointChartSnapshotKind.Pie ||
            chartKind == PowerPointChartSnapshotKind.Doughnut;

        private static double? ReadSeriesStrokeWidth(OpenXmlCompositeElement seriesElement) {
            C.ChartShapeProperties? properties = seriesElement.GetFirstChild<C.ChartShapeProperties>();
            long? widthEmus = properties?.GetFirstChild<A.Outline>()?.Width?.Value;
            return widthEmus.HasValue && widthEmus.Value > 0L
                ? PowerPointUnits.ToPoints(widthEmus.Value)
                : null;
        }

        private static OfficeColor? ReadSeriesStrokeColor(
            OpenXmlCompositeElement seriesElement, A.ColorScheme? colorScheme) {
            C.ChartShapeProperties? properties =
                seriesElement.GetFirstChild<C.ChartShapeProperties>();
            return OfficeOpenXmlThemeColorResolver.ResolveColor(
                properties?.GetFirstChild<A.Outline>()?.GetFirstChild<A.SolidFill>(),
                colorScheme);
        }

        private static bool IsSeriesStrokeVisible(OpenXmlCompositeElement seriesElement) {
            C.ChartShapeProperties? properties =
                seriesElement.GetFirstChild<C.ChartShapeProperties>();
            return properties?.GetFirstChild<A.Outline>()?.GetFirstChild<A.NoFill>() == null;
        }

        private static string? ReadTitle(C.Chart chart) {
            C.ChartText? chartText =
                chart.GetFirstChild<C.Title>()?.GetFirstChild<C.ChartText>();
            return ReadChartText(chartText);
        }

        private static string? ReadChartText(C.ChartText? chartText) {
            if (chartText == null) {
                return null;
            }

            C.RichText? richText = chartText.GetFirstChild<C.RichText>();
            string text = richText != null
                ? string.Join(Environment.NewLine,
                    richText.Elements<A.Paragraph>().Select(ReadChartParagraphText))
                : string.Concat(chartText.Descendants<A.Text>()
                    .Select(item => item.Text));
            if (!string.IsNullOrWhiteSpace(text)) {
                return text.Trim();
            }

            IReadOnlyList<string> cached = ReadCachedStrings(chartText);
            return cached.Count > 0 && !string.IsNullOrWhiteSpace(cached[0]) ? cached[0].Trim() : null;
        }

        private static string ReadChartParagraphText(A.Paragraph paragraph) {
            var builder = new System.Text.StringBuilder();
            foreach (OpenXmlElement child in paragraph.ChildElements) {
                if (child is A.Break) {
                    builder.Append(Environment.NewLine);
                } else {
                    foreach (A.Text text in child.Descendants<A.Text>()) {
                        builder.Append(text.Text);
                    }
                }
            }
            return builder.ToString();
        }

        private static string ReadSeriesName(OpenXmlElement seriesElement) {
            C.SeriesText? seriesText = seriesElement.GetFirstChild<C.SeriesText>();
            if (seriesText == null) {
                return string.Empty;
            }

            IReadOnlyList<string> cached = ReadCachedStrings(seriesText);
            if (cached.Count > 0) {
                return cached[0] ?? string.Empty;
            }

            string richText = string.Concat(seriesText.Descendants<A.Text>().Select(item => item.Text));
            return richText.Trim();
        }

        private static IReadOnlyList<string> ReadCachedStrings(OpenXmlElement? container) {
            if (container == null) {
                return Array.Empty<string>();
            }

            List<C.StringPoint> stringPoints = GetBoundedCachedPoints(container.Descendants<C.StringPoint>());
            stringPoints.Sort((left, right) => (left.Index?.Value ?? 0U).CompareTo(right.Index?.Value ?? 0U));
            if (stringPoints.Count > 0) {
                return CreateIndexedCache(
                    container,
                    stringPoints,
                    point => point.Index?.Value,
                    point => point.NumericValue?.Text ?? string.Empty,
                    string.Empty);
            }

            List<C.NumericPoint> numericPoints = GetBoundedCachedPoints(container.Descendants<C.NumericPoint>());
            numericPoints.Sort((left, right) => (left.Index?.Value ?? 0U).CompareTo(right.Index?.Value ?? 0U));
            if (numericPoints.Count > 0) {
                return CreateIndexedCache(
                    container,
                    numericPoints,
                    point => point.Index?.Value,
                    point => point.NumericValue?.Text ?? string.Empty,
                    string.Empty);
            }

            return Array.Empty<string>();
        }

        private static IReadOnlyList<double> ReadCachedNumbers(OpenXmlElement? container) {
            if (container == null) {
                return Array.Empty<double>();
            }

            List<C.NumericPoint> points = GetBoundedCachedPoints(container.Descendants<C.NumericPoint>());
            points.Sort((left, right) => (left.Index?.Value ?? 0U).CompareTo(right.Index?.Value ?? 0U));
            if (points.Count == 0) {
                return Array.Empty<double>();
            }

            return CreateIndexedCache(
                container,
                points,
                point => point.Index?.Value,
                point => {
                string? text = point.NumericValue?.Text;
                if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) &&
                    !double.IsNaN(value) &&
                    !double.IsInfinity(value)) {
                    return value;
                }

                return 0D;
                },
                0D);
        }

        private static IReadOnlyList<TValue> CreateIndexedCache<TPoint, TValue>(
            OpenXmlElement container,
            IReadOnlyList<TPoint> points,
            Func<TPoint, uint?> getIndex,
            Func<TPoint, TValue> getValue,
            TValue defaultValue) {
            int length = GetCachedPointLength(container, points, getIndex);
            var values = Enumerable.Repeat(defaultValue, length).ToArray();
            for (int i = 0; i < points.Count; i++) {
                TPoint point = points[i];
                uint? rawIndex = getIndex(point);
                int index = rawIndex.HasValue && rawIndex.Value <= int.MaxValue
                    ? (int)rawIndex.Value
                    : i;
                if (index >= 0 && index < values.Length) {
                    values[index] = getValue(point);
                }
            }

            return values;
        }

        private static List<TPoint> GetBoundedCachedPoints<TPoint>(IEnumerable<TPoint> points) {
            List<TPoint> boundedPoints = points
                .Take(PowerPointUtils.MaximumSharedChartPoints + 1).ToList();
            if (boundedPoints.Count > PowerPointUtils.MaximumSharedChartPoints) {
                throw new InvalidDataException($"The chart cache exceeds the supported limit of {PowerPointUtils.MaximumSharedChartPoints} points.");
            }

            return boundedPoints;
        }

        private static int GetCachedPointLength<TPoint>(OpenXmlElement container, IReadOnlyList<TPoint> points, Func<TPoint, uint?> getIndex) {
            if (points.Count > PowerPointUtils.MaximumSharedChartPoints) {
                throw new InvalidDataException($"The chart cache exceeds the supported limit of {PowerPointUtils.MaximumSharedChartPoints} points.");
            }

            uint? pointCount = container.Descendants<C.PointCount>().FirstOrDefault()?.Val?.Value;
            if (pointCount > PowerPointUtils.MaximumSharedChartPoints) {
                throw new InvalidDataException($"The chart cache declares more than the supported limit of {PowerPointUtils.MaximumSharedChartPoints} points.");
            }

            uint maxIndex = 0U;
            bool hasIndexedPoint = false;
            for (int i = 0; i < points.Count; i++) {
                uint? index = getIndex(points[i]);
                if (!index.HasValue) {
                    continue;
                }

                if (index.Value >= PowerPointUtils.MaximumSharedChartPoints) {
                    throw new InvalidDataException($"The chart cache point index exceeds the supported limit of {PowerPointUtils.MaximumSharedChartPoints} points.");
                }

                hasIndexedPoint = true;
                if (index.Value > maxIndex) {
                    maxIndex = index.Value;
                }
            }

            uint indexedLength = hasIndexedPoint ? maxIndex + 1U : (uint)points.Count;
            uint length = Math.Max(pointCount ?? 0U, indexedLength);
            return (int)length;
        }

        private static IReadOnlyList<string> CreateFallbackCategories(int count) {
            if (count <= 0) {
                return Array.Empty<string>();
            }

            var categories = new List<string>(count);
            for (int i = 0; i < count; i++) {
                categories.Add("Category " + (i + 1).ToString(CultureInfo.InvariantCulture));
            }

            return categories;
        }

        private static IReadOnlyList<double> NormalizeValues(IReadOnlyList<double> values, int count) {
            if (count <= 0 || values.Count == 0) {
                return Array.Empty<double>();
            }

            var normalized = new double[count];
            int take = Math.Min(values.Count, count);
            for (int i = 0; i < take; i++) {
                normalized[i] = values[i];
            }

            return normalized;
        }
    }
}
