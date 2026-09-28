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
            if (!TryReadSnapshot(colorScheme, forDataUpdate, out snapshot)) return false;
            if (!forDataUpdate && snapshot.Data.Series.Any(series => series.HasUnsupportedSharedAppearance)) {
                snapshot = null!;
                return false;
            }
            return true;
        }

        private bool TryReadSnapshot(A.ColorScheme? colorScheme,
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

                if (TryCreateAdvancedChartSnapshot(chart, plotArea, colorScheme, forDataUpdate,
                        out snapshot)) return true;

                if (TryCreateMixedChartSnapshot(chart, plotArea, colorScheme, forDataUpdate, out snapshot)) {
                    return true;
                }

                if (CountSupportedChartElements(plotArea) > 1) {
                    snapshot = null!;
                    return false;
                }

                if (plotArea.GetFirstChild<C.BarChart>() is C.BarChart barChart) {
                    PowerPointChartSnapshotKind kind = GetBarChartSnapshotKind(barChart);
                    PowerPointChartData? data = ReadCategorySeriesData(barChart.Elements<C.BarChartSeries>().Cast<OpenXmlCompositeElement>(), kind, colorScheme, forDataUpdate: forDataUpdate);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, kind, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.LineChart>() is C.LineChart lineChart) {
                    PowerPointChartSnapshotKind kind = GetLineChartSnapshotKind(lineChart);
                    PowerPointChartData? data = ReadCategorySeriesData(lineChart.Elements<C.LineChartSeries>().Cast<OpenXmlCompositeElement>(), kind, colorScheme, forDataUpdate: forDataUpdate);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, kind, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.AreaChart>() is C.AreaChart areaChart) {
                    PowerPointChartSnapshotKind kind = GetAreaChartSnapshotKind(areaChart);
                    PowerPointChartData? data = ReadCategorySeriesData(areaChart.Elements<C.AreaChartSeries>().Cast<OpenXmlCompositeElement>(), kind, colorScheme, forDataUpdate: forDataUpdate);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, kind, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.RadarChart>() is C.RadarChart radarChart) {
                    PowerPointChartData? data = ReadCategorySeriesData(radarChart.Elements<C.RadarChartSeries>().Cast<OpenXmlCompositeElement>(), PowerPointChartSnapshotKind.Radar, colorScheme, forDataUpdate: forDataUpdate);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, PowerPointChartSnapshotKind.Radar, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.ScatterChart>() is C.ScatterChart scatterChart) {
                    PowerPointChartData? data = ReadScatterSeriesData(scatterChart.Elements<C.ScatterChartSeries>(), colorScheme, forDataUpdate: forDataUpdate);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, PowerPointChartSnapshotKind.Scatter, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.BubbleChart>() is C.BubbleChart bubbleChart) {
                    if (!forDataUpdate && OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartSeriesReader.HasUnsupportedBubblePresentation(
                        chartPart, chart, plotArea, bubbleChart, PowerPointUtils.MaximumSharedChartPoints)) {
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
                    PowerPointChartData? data = ReadCategorySeriesData(pieChart.Elements<C.PieChartSeries>().Cast<OpenXmlCompositeElement>(), PowerPointChartSnapshotKind.Pie, colorScheme, forDataUpdate: forDataUpdate);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, PowerPointChartSnapshotKind.Pie, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.DoughnutChart>() is C.DoughnutChart doughnutChart) {
                    PowerPointChartData? data = ReadCategorySeriesData(doughnutChart.Elements<C.PieChartSeries>().Cast<OpenXmlCompositeElement>(), PowerPointChartSnapshotKind.Doughnut, colorScheme, forDataUpdate: forDataUpdate);
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

        private static int CountSupportedChartElements(C.PlotArea plotArea) {
            return plotArea.Elements<C.BarChart>().Count()
                + plotArea.Elements<C.LineChart>().Count()
                + plotArea.Elements<C.AreaChart>().Count()
                + plotArea.Elements<C.RadarChart>().Count()
                + plotArea.Elements<C.ScatterChart>().Count()
                + plotArea.Elements<C.BubbleChart>().Count()
                + plotArea.Elements<C.PieChart>().Count()
                + plotArea.Elements<C.DoughnutChart>().Count()
                + plotArea.ChildElements.Count(element =>
                    AdvancedChartProjections.ContainsKey(element.LocalName));
        }

        private bool TryCreateAdvancedChartSnapshot(C.Chart chart,
            C.PlotArea plotArea, A.ColorScheme? colorScheme, bool forDataUpdate,
            out PowerPointChartSnapshot snapshot) {
            snapshot = null!;
            OpenXmlCompositeElement[] groups = plotArea.ChildElements
                .OfType<OpenXmlCompositeElement>()
                .Where(element => AdvancedChartProjections.ContainsKey(element.LocalName))
                .ToArray();
            if (groups.Length != 1 || CountSupportedChartElements(plotArea) != 1) return false;
            PowerPointChartSnapshotKind kind = MapAdvancedSnapshotKind(
                GetAdvancedProjection(groups[0]));
            PowerPointChartData? data = ReadCategorySeriesData(groups[0]
                .ChildElements.OfType<OpenXmlCompositeElement>()
                .Where(element => element.LocalName == "ser"), kind, colorScheme, forDataUpdate: forDataUpdate);
            if (data == null) return false;
            snapshot = CreateSnapshot(chart, kind, data);
            return true;
        }

        private static PowerPointChartSnapshotKind MapAdvancedSnapshotKind(
            OfficeChartKind kind) => kind switch {
                OfficeChartKind.ColumnClustered => PowerPointChartSnapshotKind.ClusteredColumn,
                OfficeChartKind.ColumnStacked => PowerPointChartSnapshotKind.StackedColumn,
                OfficeChartKind.ColumnStacked100 => PowerPointChartSnapshotKind.StackedColumn100,
                OfficeChartKind.BarClustered => PowerPointChartSnapshotKind.ClusteredBar,
                OfficeChartKind.BarStacked => PowerPointChartSnapshotKind.StackedBar,
                OfficeChartKind.BarStacked100 => PowerPointChartSnapshotKind.StackedBar100,
                OfficeChartKind.Area => PowerPointChartSnapshotKind.Area,
                OfficeChartKind.AreaStacked => PowerPointChartSnapshotKind.StackedArea,
                OfficeChartKind.AreaStacked100 => PowerPointChartSnapshotKind.StackedArea100,
                OfficeChartKind.LineStacked => PowerPointChartSnapshotKind.StackedLine,
                OfficeChartKind.LineStacked100 => PowerPointChartSnapshotKind.StackedLine100,
                OfficeChartKind.Pie => PowerPointChartSnapshotKind.Pie,
                _ => PowerPointChartSnapshotKind.Line
            };

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
                element is not C.DoughnutChart &&
                !AdvancedChartProjections.ContainsKey(element.LocalName));

        private bool TryCreateMixedChartSnapshot(C.Chart chart, C.PlotArea plotArea, A.ColorScheme? colorScheme,
            bool forDataUpdate, out PowerPointChartSnapshot snapshot) {
            snapshot = null!;
            int supportedGroupCount = CountSupportedChartElements(plotArea);
            if (supportedGroupCount <= 1 || plotArea.ChildElements.Any(element =>
                    AdvancedChartProjections.ContainsKey(element.LocalName))) {
                return false;
            }
            if (plotArea.Elements<C.BubbleChart>().Any()) {
                return false;
            }

            var parts = new List<(PowerPointChartSnapshotKind Kind, PowerPointChartData Data, bool HasSourceCategories)>();
            OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartSeriesReader.ValidatePlotBudget(plotArea, PowerPointUtils.MaximumSharedChartPoints);
            var axisGroups = OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartAxisGroups.Create(plotArea);
            var referencedAxes = new Dictionary<OfficeChartAxisGroup, (uint Category, uint Value)>();
            foreach (OpenXmlCompositeElement layer in plotArea.ChildElements.OfType<OpenXmlCompositeElement>()
                         .Where(element => element is C.BarChart or C.LineChart or C.AreaChart)) {
                OpenXmlCompositeElement?[] axes = layer.Elements<C.AxisId>()
                    .Select(reference => axisGroups.Resolve(reference.Val?.Value)).ToArray();
                if (axes.Length != 2 || axes.OfType<C.CategoryAxis>().FirstOrDefault() is not C.CategoryAxis category ||
                    axes.OfType<C.ValueAxis>().FirstOrDefault() is not C.ValueAxis value ||
                    category.AxisId?.Val?.Value is not uint categoryId || value.AxisId?.Val?.Value is not uint valueId)
                    return false;
                OfficeChartAxisGroup group = axisGroups.Read(layer);
                if (referencedAxes.TryGetValue(group, out var prior) &&
                    (prior.Category != categoryId || prior.Value != valueId)) return false;
                referencedAxes[group] = (categoryId, valueId);
            }
            var projectionBudget = new OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartSeriesReader.ProjectionBudget();
            foreach (OpenXmlElement element in plotArea.ChildElements) {
                if (element is C.BarChart barChart) {
                    PowerPointChartSnapshotKind kind = GetBarChartSnapshotKind(barChart);
                    OpenXmlCompositeElement[] sourceSeries = barChart.Elements<C.BarChartSeries>()
                        .Cast<OpenXmlCompositeElement>().ToArray();
                    PowerPointChartData? data = ReadCategorySeriesData(
                        sourceSeries, kind, colorScheme,
                        axisGroups.Read(barChart), validatePlot: false, projectionBudget: projectionBudget, forDataUpdate: forDataUpdate);
                    if (data != null) {
                        parts.Add((kind, data, HasSourceCategoryPoints(sourceSeries)));
                    }
                } else if (element is C.LineChart lineChart) {
                    PowerPointChartSnapshotKind kind = GetLineChartSnapshotKind(lineChart);
                    OpenXmlCompositeElement[] sourceSeries = lineChart.Elements<C.LineChartSeries>()
                        .Cast<OpenXmlCompositeElement>().ToArray();
                    PowerPointChartData? data = ReadCategorySeriesData(
                        sourceSeries, kind, colorScheme,
                        axisGroups.Read(lineChart), validatePlot: false, projectionBudget: projectionBudget, forDataUpdate: forDataUpdate);
                    if (data != null) {
                        parts.Add((kind, data, HasSourceCategoryPoints(sourceSeries)));
                    }
                } else if (element is C.AreaChart areaChart) {
                    PowerPointChartSnapshotKind kind = GetAreaChartSnapshotKind(areaChart);
                    OpenXmlCompositeElement[] sourceSeries = areaChart.Elements<C.AreaChartSeries>()
                        .Cast<OpenXmlCompositeElement>().ToArray();
                    PowerPointChartData? data = ReadCategorySeriesData(
                        sourceSeries, kind, colorScheme,
                        axisGroups.Read(areaChart), validatePlot: false, projectionBudget: projectionBudget, forDataUpdate: forDataUpdate);
                    if (data != null) {
                        parts.Add((kind, data, HasSourceCategoryPoints(sourceSeries)));
                    }
                } else if (element is C.ScatterChart scatterChart) {
                    PowerPointChartData? data = ReadScatterSeriesData(scatterChart.Elements<C.ScatterChartSeries>(), colorScheme, validatePlot: false, projectionBudget: projectionBudget, forDataUpdate: forDataUpdate);
                    if (data != null) {
                        parts.Add((PowerPointChartSnapshotKind.Scatter, data, false));
                    }
                }
            }

            if (parts.Count <= 1 || parts.Count != supportedGroupCount) {
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

            int sourceCategoryPart = parts.FindIndex(part => part.HasSourceCategories);
            IReadOnlyList<string> categories = parts[sourceCategoryPart >= 0 ? sourceCategoryPart : 0].Data.Categories;
            if (parts[0].Kind != PowerPointChartSnapshotKind.Scatter &&
                parts.Any(part => part.HasSourceCategories &&
                    !part.Data.Categories.SequenceEqual(categories, StringComparer.Ordinal))) return false;
            var series = new List<PowerPointChartSeries>();
            foreach (var part in parts) {
                foreach (PowerPointChartSeries item in part.Data.Series) {
                    if (item.Values.Count != categories.Count && !HasAlignedScatterPoints(item)) return false;
                    series.Add(item);
                }
            }

            if (series.Count == 0) {
                return false;
            }

            if (series.Where(item => item.SourceOrder.HasValue).GroupBy(item => item.SourceOrder).Any(group => group.Count() > 1)) return false;
            snapshot = CreateSnapshot(chart, parts[0].Kind, new PowerPointChartData(categories,
                series.OrderBy(item => item.SourceOrder ?? uint.MaxValue)));
            return true;
        }

        private static bool HasAlignedScatterPoints(PowerPointChartSeries series) =>
            series.XValues != null &&
            series.XValues.Count == series.Values.Count &&
            series.Values.Count > 0;

        private static bool HasSourceCategoryPoints(IEnumerable<OpenXmlCompositeElement> series) =>
            series.Any(item => item.GetFirstChild<C.CategoryAxisData>() is { } categories &&
                categories.Descendants().Any(point => point is C.StringPoint or C.NumericPoint));

        private static bool IsHorizontalBarKind(PowerPointChartSnapshotKind kind) =>
            kind == PowerPointChartSnapshotKind.ClusteredBar ||
            kind == PowerPointChartSnapshotKind.StackedBar ||
            kind == PowerPointChartSnapshotKind.StackedBar100;

        private PowerPointChartSnapshot CreateSnapshot(C.Chart chart,
            PowerPointChartSnapshotKind kind, PowerPointChartData data,
            OfficeChartBubbleSizeMode bubbleSizeMode = OfficeChartBubbleSizeMode.Area,
            double bubbleScalePercent = 100D) {
            HashSet<uint> hiddenLegendSeries = GetHiddenLegendSeriesIndexes(chart);
            bool hasLegend = chart.GetFirstChild<C.Legend>() != null;
            for (int seriesIndex = 0; seriesIndex < data.Series.Count; seriesIndex++) {
                PowerPointChartSeries series = data.Series[seriesIndex];
                uint sourceIndex = series.SourceIndex ?? (uint)seriesIndex;
                uint legendIndex = kind is PowerPointChartSnapshotKind.Pie or PowerPointChartSnapshotKind.Doughnut
                    ? sourceIndex : (uint)seriesIndex;
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
                ReadSharedTextStyle(chart),
                OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartRadialLayout.Read(chart));
        }

        private OfficeChartLayout ReadChartLayout(
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
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (kind == PowerPointChartSnapshotKind.Bubble &&
                plotArea != null &&
                plotArea.GetFirstChild<C.BubbleChart>() is C.BubbleChart bubble &&
                OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartSeriesReader.TryGetReferencedBubbleAxes(
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
            } else if (plotArea != null) {
                OpenXmlCompositeElement? categoryAxis =
                    (OpenXmlCompositeElement?)plotArea
                        .Elements<C.CategoryAxis>().FirstOrDefault()
                    ?? plotArea.Elements<C.DateAxis>().FirstOrDefault();
                C.ValueAxis? valueAxis = plotArea.Elements<C.ValueAxis>()
                    .FirstOrDefault();
                if (categoryAxis != null) {
                    horizontalAxisTitle = ReadAxisTitle(categoryAxis);
                    verticalAxisTitle = valueAxis == null
                        ? null
                        : ReadAxisTitle(valueAxis);
                } else {
                    C.ValueAxis? positionedHorizontalAxis = plotArea
                        .Elements<C.ValueAxis>().FirstOrDefault(axis =>
                            axis.AxisPosition?.Val?.Value
                                == C.AxisPositionValues.Bottom
                            || axis.AxisPosition?.Val?.Value
                                == C.AxisPositionValues.Top);
                    C.ValueAxis? positionedVerticalAxis = plotArea
                        .Elements<C.ValueAxis>().FirstOrDefault(axis =>
                            axis.AxisPosition?.Val?.Value
                                == C.AxisPositionValues.Left
                            || axis.AxisPosition?.Val?.Value
                                == C.AxisPositionValues.Right);
                    horizontalAxisTitle = positionedHorizontalAxis == null
                        ? null
                        : ReadAxisTitle(positionedHorizontalAxis);
                    verticalAxisTitle = positionedVerticalAxis == null
                        ? null
                        : ReadAxisTitle(positionedVerticalAxis);
                }
            }

            TryReadAxisTitleTypeface(chart,
                ReadChartDefaultTypeface(chart), out string? axisTitleFont);

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
                verticalAxisMinorTickMark: verticalMinorTickMark,
                axisTitleFontFamily: axisTitleFont);
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

        private static string? ReadAxisTitle(OpenXmlCompositeElement axis) =>
            ReadChartText(
                axis.GetFirstChild<C.Title>()?.GetFirstChild<C.ChartText>());

        private static string? ReadAxisNumberFormat(C.ValueAxis axis) {
            string? format = axis.GetFirstChild<C.NumberingFormat>()?
                .FormatCode?.Value;
            return string.IsNullOrWhiteSpace(format) ? null : format;
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
            OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary, bool validatePlot = true,
            OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartSeriesReader.ProjectionBudget? projectionBudget = null,
            bool forDataUpdate = false) =>
            ProjectSharedSeries(OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartSeriesReader.ReadCategories(seriesElements,
                PowerPointChartSnapshotMapper.MapKind(chartKind ?? PowerPointChartSnapshotKind.ClusteredColumn),
                colorScheme, axisGroup, PowerPointUtils.MaximumSharedChartPoints, validatePlot, projectionBudget,
                maximumPointOverrides: PowerPointUtils.MaximumSharedChartPoints, forDataUpdate: forDataUpdate), chartKind);

        private static PowerPointChartData? ReadScatterSeriesData(IEnumerable<C.ScatterChartSeries> seriesElements,
            A.ColorScheme? colorScheme = null, bool validatePlot = true,
            OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartSeriesReader.ProjectionBudget? projectionBudget = null,
            bool forDataUpdate = false) =>
            ProjectSharedSeries(OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartSeriesReader.ReadScatter(seriesElements,
                colorScheme, PowerPointUtils.MaximumSharedChartPoints, validatePlot, projectionBudget,
                maximumPointOverrides: PowerPointUtils.MaximumSharedChartPoints, forDataUpdate: forDataUpdate), PowerPointChartSnapshotKind.Scatter);

        private static PowerPointChartData? ProjectSharedSeries(OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartSeriesReader.Result? source,
            PowerPointChartSnapshotKind? kind) => source == null ? null : new PowerPointChartData(source.Categories,
                source.Series.Select(item => new PowerPointChartSeries(item.Data.Name, item.Data.Values, item.Data.XValues,
                    kind, item.Data.Color, item.Data.BubbleSizes != null ? item.Data.MarkerOutlineWidth : item.Data.StrokeWidth, item.Data.AxisGroup) {
                    BubbleSizes = item.Data.BubbleSizes,
                    StrokeColor = item.Data.BubbleSizes != null ? item.Data.MarkerOutlineColor : null,
                    ShowStroke = item.Data.ShowMarkerOutline,
                    PointColors = item.Data.PointColors,
                    PointStyles = item.Data.PointStyles,
                    SourceIndex = item.SourceIndex,
                    SourceOrder = item.SourceOrder,
                    SharedAppearance = item.Data,
                    HasUnsupportedSharedAppearance = item.HasUnsupportedAppearance
                }));

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

        private static IReadOnlyList<string> ReadCachedStrings(OpenXmlElement? container) =>
            OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartCacheReader.ReadCachedStrings(container, PowerPointUtils.MaximumSharedChartPoints);

        private static IReadOnlyList<double> ReadCachedNumbers(OpenXmlElement? container) =>
            OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartCacheReader.ReadCachedNumbers(container, PowerPointUtils.MaximumSharedChartPoints);
        private static List<TPoint> GetBoundedCachedPoints<TPoint>(IEnumerable<TPoint> points) =>
            OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartCacheReader.GetBoundedCachedPoints(points, PowerPointUtils.MaximumSharedChartPoints);

        private static int GetCachedPointLength<TPoint>(OpenXmlElement container, IReadOnlyList<TPoint> points, Func<TPoint, uint?> getIndex) =>
            OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartCacheReader.GetCachedPointLength(container, points, getIndex, PowerPointUtils.MaximumSharedChartPoints);

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
