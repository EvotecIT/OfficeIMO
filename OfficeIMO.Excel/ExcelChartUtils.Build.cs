using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using ChartIndex = DocumentFormat.OpenXml.Drawing.Charts.Index;

namespace OfficeIMO.Excel {
    internal static partial class ExcelChartUtils {
        internal static void PopulateChart(ChartPart chartPart, ExcelChartType type, ExcelChartDataRange range,
            ExcelChartData? data, string? title = null) {
            if (chartPart == null) {
                throw new ArgumentNullException(nameof(chartPart));
            }

            ChartSpace chartSpace = new();
            chartSpace.AddNamespaceDeclaration("c", ChartNamespace);
            chartSpace.AddNamespaceDeclaration("a", DrawingNamespace);
            chartSpace.AddNamespaceDeclaration("r", RelationshipNamespace);

            chartSpace.Append(new Date1904 { Val = false });
            chartSpace.Append(new EditingLanguage { Val = "en-US" });

            Chart chart = new();
            if (!string.IsNullOrWhiteSpace(title)) {
                chart.Append(CreateChartTitle(title!));
            }
            chart.Append(new AutoTitleDeleted { Val = string.IsNullOrWhiteSpace(title) });

            PlotArea plotArea = new() { Layout = new Layout() };
            List<SeriesDescriptor> descriptors = BuildSeriesDescriptors(range, data, type);
            if (descriptors.Any(d => d.ChartType == ExcelChartType.Bubble)) {
                throw new NotSupportedException("Bubble charts require explicit X/Y/size ranges. Use AddBubbleChartFromRanges.");
            }
            bool hasSecondary = descriptors.Any(d => d.AxisGroup == ExcelChartAxisGroup.Secondary);
            bool hasScatter = descriptors.Any(d => d.ChartType == ExcelChartType.Scatter);
            if (hasScatter && descriptors.Any(d => d.ChartType != ExcelChartType.Scatter)) {
                throw new NotSupportedException("Scatter charts cannot be combined with other chart types.");
            }
            bool hasMultipleTypes = descriptors.Select(d => d.ChartType).Distinct().Count() > 1;

            if (hasScatter) {
                if (hasMultipleTypes) {
                    throw new NotSupportedException("Scatter charts cannot be combined with other chart types.");
                }

                uint xAxisId = ExcelChartAxisIdGenerator.GetNextId();
                uint yAxisId = ExcelChartAxisIdGenerator.GetNextId();
                plotArea.Append(CreateScatterChart(range, descriptors, xAxisId, yAxisId, data));
                plotArea.Append(CreateValueAxis(xAxisId, yAxisId, AxisPositionValues.Bottom));
                plotArea.Append(CreateValueAxis(yAxisId, xAxisId, AxisPositionValues.Left));
            } else if (hasMultipleTypes || hasSecondary) {
                BuildComboPlotArea(plotArea, range, descriptors);
            } else {
                ExcelChartType chartType = descriptors.Count > 0 ? descriptors[0].ChartType : type;
                uint categoryAxisId = ExcelChartAxisIdGenerator.GetNextId();
                uint valueAxisId = ExcelChartAxisIdGenerator.GetNextId();

                switch (chartType) {
                    case ExcelChartType.ColumnClustered:
                        plotArea.Append(CreateBarChart(range, descriptors, BarDirectionValues.Column, BarGroupingValues.Clustered, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId));
                        break;
                    case ExcelChartType.ColumnStacked:
                        plotArea.Append(CreateBarChart(range, descriptors, BarDirectionValues.Column, BarGroupingValues.Stacked, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId));
                        break;
                    case ExcelChartType.BarClustered:
                        plotArea.Append(CreateBarChart(range, descriptors, BarDirectionValues.Bar, BarGroupingValues.Clustered, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId, AxisPositionValues.Left));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId, AxisPositionValues.Bottom));
                        break;
                    case ExcelChartType.BarStacked:
                        plotArea.Append(CreateBarChart(range, descriptors, BarDirectionValues.Bar, BarGroupingValues.Stacked, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId, AxisPositionValues.Left));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId, AxisPositionValues.Bottom));
                        break;
                    case ExcelChartType.Line:
                        plotArea.Append(CreateLineChart(range, descriptors, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId));
                        break;
                    case ExcelChartType.Area:
                        plotArea.Append(CreateAreaChart(range, descriptors, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId));
                        break;
                    case ExcelChartType.Pie:
                        plotArea.Append(CreatePieChart(range, descriptors));
                        break;
                    case ExcelChartType.Doughnut:
                        plotArea.Append(CreateDoughnutChart(range, descriptors));
                        break;
                    default:
                        plotArea.Append(CreateBarChart(range, descriptors, BarDirectionValues.Column, BarGroupingValues.Clustered, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId));
                        break;
                }
            }

            chart.Append(plotArea);
            chart.Append(new Legend(
                new LegendPosition { Val = LegendPositionValues.Bottom },
                new Layout(),
                new Overlay { Val = false }));
            chart.Append(new PlotVisibleOnly { Val = true });
            chart.Append(new DisplayBlanksAs { Val = DisplayBlanksAsValues.Gap });
            chart.Append(new ShowDataLabelsOverMaximum { Val = false });

            chartSpace.Append(chart);
            chartPart.ChartSpace = chartSpace;
        }

        internal static void PopulateChartFromSeriesRanges(ChartPart chartPart, ExcelChartType type, string defaultSheetName,
            IReadOnlyList<ExcelChartSeriesRange> seriesRanges, string? title = null) {
            if (chartPart == null) {
                throw new ArgumentNullException(nameof(chartPart));
            }
            if (string.IsNullOrWhiteSpace(defaultSheetName)) {
                throw new ArgumentException("Default sheet name cannot be null or empty.", nameof(defaultSheetName));
            }
            if (seriesRanges == null) {
                throw new ArgumentNullException(nameof(seriesRanges));
            }
            if (seriesRanges.Count == 0) {
                throw new ArgumentException("At least one series range is required.", nameof(seriesRanges));
            }
            if (type != ExcelChartType.Scatter && type != ExcelChartType.Bubble) {
                throw new NotSupportedException("Only scatter and bubble charts support explicit X/Y range definitions.");
            }
            if (type == ExcelChartType.Bubble && seriesRanges.Any(r => string.IsNullOrWhiteSpace(r.BubbleSizeRangeA1))) {
                throw new ArgumentException("Bubble charts require bubble size ranges for each series.", nameof(seriesRanges));
            }

            ChartSpace chartSpace = new();
            chartSpace.AddNamespaceDeclaration("c", ChartNamespace);
            chartSpace.AddNamespaceDeclaration("a", DrawingNamespace);
            chartSpace.AddNamespaceDeclaration("r", RelationshipNamespace);

            chartSpace.Append(new Date1904 { Val = false });
            chartSpace.Append(new EditingLanguage { Val = "en-US" });

            Chart chart = new();
            if (!string.IsNullOrWhiteSpace(title)) {
                chart.Append(CreateChartTitle(title!));
            }
            chart.Append(new AutoTitleDeleted { Val = string.IsNullOrWhiteSpace(title) });

            PlotArea plotArea = new() { Layout = new Layout() };
            uint xAxisId = ExcelChartAxisIdGenerator.GetNextId();
            uint yAxisId = ExcelChartAxisIdGenerator.GetNextId();

            if (type == ExcelChartType.Scatter) {
                plotArea.Append(CreateScatterChartFromRanges(seriesRanges, defaultSheetName, xAxisId, yAxisId));
            } else {
                plotArea.Append(CreateBubbleChartFromRanges(seriesRanges, defaultSheetName, xAxisId, yAxisId));
            }

            plotArea.Append(CreateValueAxis(xAxisId, yAxisId, AxisPositionValues.Bottom));
            plotArea.Append(CreateValueAxis(yAxisId, xAxisId, AxisPositionValues.Left));

            chart.Append(plotArea);
            chart.Append(new Legend(
                new LegendPosition { Val = LegendPositionValues.Bottom },
                new Layout(),
                new Overlay { Val = false }));
            chart.Append(new PlotVisibleOnly { Val = true });
            chart.Append(new DisplayBlanksAs { Val = DisplayBlanksAsValues.Gap });
            chart.Append(new ShowDataLabelsOverMaximum { Val = false });

            chartSpace.Append(chart);
            chartPart.ChartSpace = chartSpace;
        }

        private static void BuildComboPlotArea(PlotArea plotArea, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            bool hasSecondary = descriptors.Any(d => d.AxisGroup == ExcelChartAxisGroup.Secondary);
            bool hasScatter = descriptors.Any(d => d.ChartType == ExcelChartType.Scatter);
            if (descriptors.Any(d => d.ChartType == ExcelChartType.Bubble)) {
                throw new NotSupportedException("Bubble charts cannot be combined with other chart types.");
            }
            if (hasScatter && descriptors.Any(d => d.ChartType != ExcelChartType.Scatter)) {
                throw new NotSupportedException("Scatter charts cannot be combined with other chart types.");
            }

            bool hasBar = descriptors.Any(d => IsBarChartType(d.ChartType));
            bool hasNonBar = descriptors.Any(d => !IsBarChartType(d.ChartType));
            if (hasBar && hasNonBar) {
                throw new NotSupportedException("Cannot combine horizontal bar charts with other chart types.");
            }

            if (descriptors.Any(d => d.ChartType == ExcelChartType.Pie || d.ChartType == ExcelChartType.Doughnut)) {
                throw new NotSupportedException("Pie and doughnut charts cannot be combined with other chart types.");
            }

            bool isBarOrientation = hasBar;
            AxisPositionValues primaryCategoryPosition = isBarOrientation ? AxisPositionValues.Left : AxisPositionValues.Bottom;
            AxisPositionValues primaryValuePosition = isBarOrientation ? AxisPositionValues.Bottom : AxisPositionValues.Left;
            AxisPositionValues secondaryCategoryPosition = isBarOrientation ? AxisPositionValues.Right : AxisPositionValues.Top;
            AxisPositionValues secondaryValuePosition = isBarOrientation ? AxisPositionValues.Top : AxisPositionValues.Right;

            uint primaryCategoryId = ExcelChartAxisIdGenerator.GetNextId();
            uint primaryValueId = ExcelChartAxisIdGenerator.GetNextId();
            uint secondaryCategoryId = hasSecondary ? ExcelChartAxisIdGenerator.GetNextId() : 0;
            uint secondaryValueId = hasSecondary ? ExcelChartAxisIdGenerator.GetNextId() : 0;

            foreach (var group in descriptors.GroupBy(d => new { d.ChartType, d.AxisGroup })) {
                uint categoryAxisId = group.Key.AxisGroup == ExcelChartAxisGroup.Secondary ? secondaryCategoryId : primaryCategoryId;
                uint valueAxisId = group.Key.AxisGroup == ExcelChartAxisGroup.Secondary ? secondaryValueId : primaryValueId;
                var groupDescriptors = group.ToList();

                switch (group.Key.ChartType) {
                    case ExcelChartType.ColumnClustered:
                    case ExcelChartType.ColumnStacked:
                    case ExcelChartType.BarClustered:
                    case ExcelChartType.BarStacked: {
                        var settings = GetBarChartSettings(group.Key.ChartType);
                        plotArea.Append(CreateBarChart(range, groupDescriptors, settings.Direction, settings.Grouping, categoryAxisId, valueAxisId));
                        break;
                    }
                    case ExcelChartType.Line:
                        plotArea.Append(CreateLineChart(range, groupDescriptors, categoryAxisId, valueAxisId));
                        break;
                    case ExcelChartType.Area:
                        plotArea.Append(CreateAreaChart(range, groupDescriptors, categoryAxisId, valueAxisId));
                        break;
                    default:
                        throw new NotSupportedException($"Chart type {group.Key.ChartType} is not supported in combination charts.");
                }
            }

            plotArea.Append(CreateCategoryAxis(primaryCategoryId, primaryValueId, primaryCategoryPosition));
            plotArea.Append(CreateValueAxis(primaryValueId, primaryCategoryId, primaryValuePosition));
            if (hasSecondary) {
                plotArea.Append(CreateCategoryAxis(secondaryCategoryId, secondaryValueId, secondaryCategoryPosition));
                plotArea.Append(CreateValueAxis(secondaryValueId, secondaryCategoryId, secondaryValuePosition));
            }
        }

        private static BarChart CreateBarChart(ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors,
            BarDirectionValues direction, BarGroupingValues grouping, uint categoryAxisId, uint valueAxisId) {
            var barChart = new BarChart(
                new BarDirection { Val = direction },
                new BarGrouping { Val = grouping },
                new VaryColors { Val = false });

            foreach (var descriptor in seriesDescriptors) {
                barChart.Append(CreateBarChartSeries(descriptor.Index, range, descriptor.Series));
            }

            barChart.Append(CreateDefaultDataLabels());
            barChart.Append(new GapWidth { Val = (UInt16Value)219U });
            barChart.Append(new Overlap { Val = (SByteValue)(sbyte)-27 });
            barChart.Append(new AxisId { Val = categoryAxisId });
            barChart.Append(new AxisId { Val = valueAxisId });
            return barChart;
        }

        private static BarChartSeries CreateBarChartSeries(int seriesIndex, ExcelChartDataRange range, ExcelChartSeries? series) {
            return new BarChartSeries(
                new ChartIndex { Val = (uint)seriesIndex },
                new Order { Val = (uint)seriesIndex },
                CreateSeriesText(range, seriesIndex, series?.Name ?? $"Series {seriesIndex + 1}"),
                new InvertIfNegative { Val = false },
                CreateCategoryAxisData(range),
                CreateValues(range, seriesIndex, series)
            );
        }

        private static LineChart CreateLineChart(ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors,
            uint categoryAxisId, uint valueAxisId) {
            var lineChart = new LineChart(
                new Grouping { Val = GroupingValues.Standard },
                new VaryColors { Val = false });

            foreach (var descriptor in seriesDescriptors) {
                lineChart.Append(CreateLineChartSeries(descriptor.Index, range, descriptor.Series));
            }

            lineChart.Append(CreateDefaultDataLabels());
            lineChart.Append(new AxisId { Val = categoryAxisId });
            lineChart.Append(new AxisId { Val = valueAxisId });
            return lineChart;
        }

        private static LineChartSeries CreateLineChartSeries(int seriesIndex, ExcelChartDataRange range, ExcelChartSeries? series) {
            return new LineChartSeries(
                new ChartIndex { Val = (uint)seriesIndex },
                new Order { Val = (uint)seriesIndex },
                CreateSeriesText(range, seriesIndex, series?.Name ?? $"Series {seriesIndex + 1}"),
                CreateCategoryAxisData(range),
                CreateValues(range, seriesIndex, series)
            );
        }

        private static AreaChart CreateAreaChart(ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors,
            uint categoryAxisId, uint valueAxisId) {
            var areaChart = new AreaChart(
                new Grouping { Val = GroupingValues.Standard },
                new VaryColors { Val = false });

            foreach (var descriptor in seriesDescriptors) {
                areaChart.Append(CreateAreaChartSeries(descriptor.Index, range, descriptor.Series));
            }

            areaChart.Append(CreateDefaultDataLabels());
            areaChart.Append(new AxisId { Val = categoryAxisId });
            areaChart.Append(new AxisId { Val = valueAxisId });
            return areaChart;
        }

        private static AreaChartSeries CreateAreaChartSeries(int seriesIndex, ExcelChartDataRange range, ExcelChartSeries? series) {
            return new AreaChartSeries(
                new ChartIndex { Val = (uint)seriesIndex },
                new Order { Val = (uint)seriesIndex },
                CreateSeriesText(range, seriesIndex, series?.Name ?? $"Series {seriesIndex + 1}"),
                CreateCategoryAxisData(range),
                CreateValues(range, seriesIndex, series)
            );
        }

        private static PieChart CreatePieChart(ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors) {
            var pieChart = new PieChart(new VaryColors { Val = true });

            foreach (var descriptor in seriesDescriptors) {
                pieChart.Append(CreatePieChartSeries(descriptor.Index, range, descriptor.Series));
            }

            pieChart.Append(CreateDefaultDataLabels());
            return pieChart;
        }

        private static PieChartSeries CreatePieChartSeries(int seriesIndex, ExcelChartDataRange range, ExcelChartSeries? series) {
            return new PieChartSeries(
                new ChartIndex { Val = (uint)seriesIndex },
                new Order { Val = (uint)seriesIndex },
                CreateSeriesText(range, seriesIndex, series?.Name ?? $"Series {seriesIndex + 1}"),
                CreateCategoryAxisData(range),
                CreateValues(range, seriesIndex, series)
            );
        }

        private static DoughnutChart CreateDoughnutChart(ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors) {
            var chart = new DoughnutChart(
                new VaryColors { Val = true },
                new HoleSize { Val = 50 });

            foreach (var descriptor in seriesDescriptors) {
                chart.Append(CreatePieChartSeries(descriptor.Index, range, descriptor.Series));
            }

            chart.Append(CreateDefaultDataLabels());
            return chart;
        }

        private static ScatterChart CreateScatterChart(ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors,
            uint xAxisId, uint yAxisId, ExcelChartData? data) {
            var scatterChart = new ScatterChart(
                new ScatterStyle { Val = ScatterStyleValues.LineMarker },
                new VaryColors { Val = false });

            IReadOnlyList<double>? xValues = data != null ? ParseNumericCategories(data.Categories) : null;
            foreach (var descriptor in seriesDescriptors) {
                scatterChart.Append(CreateScatterChartSeries(descriptor.Index, range, descriptor.Series, xValues));
            }

            scatterChart.Append(CreateDefaultDataLabels());
            scatterChart.Append(new AxisId { Val = xAxisId });
            scatterChart.Append(new AxisId { Val = yAxisId });
            return scatterChart;
        }

        private static ScatterChartSeries CreateScatterChartSeries(int seriesIndex, ExcelChartDataRange range,
            ExcelChartSeries? series, IReadOnlyList<double>? xValues) {
            return new ScatterChartSeries(
                new ChartIndex { Val = (uint)seriesIndex },
                new Order { Val = (uint)seriesIndex },
                CreateSeriesText(range, seriesIndex, series?.Name ?? $"Series {seriesIndex + 1}"),
                CreateXValues(range, xValues),
                CreateYValues(range, seriesIndex, series)
            );
        }

        private static ScatterChart CreateScatterChartFromRanges(IReadOnlyList<ExcelChartSeriesRange> seriesRanges,
            string defaultSheetName, uint xAxisId, uint yAxisId) {
            var scatterChart = new ScatterChart(
                new ScatterStyle { Val = ScatterStyleValues.LineMarker },
                new VaryColors { Val = false });

            for (int i = 0; i < seriesRanges.Count; i++) {
                scatterChart.Append(CreateScatterChartSeriesFromRanges(i, seriesRanges[i], defaultSheetName));
            }

            scatterChart.Append(CreateDefaultDataLabels());
            scatterChart.Append(new AxisId { Val = xAxisId });
            scatterChart.Append(new AxisId { Val = yAxisId });
            return scatterChart;
        }

        private static ScatterChartSeries CreateScatterChartSeriesFromRanges(int seriesIndex, ExcelChartSeriesRange seriesRange,
            string defaultSheetName) {
            string name = string.IsNullOrWhiteSpace(seriesRange.Name) ? $"Series {seriesIndex + 1}" : seriesRange.Name;
            return new ScatterChartSeries(
                new ChartIndex { Val = (uint)seriesIndex },
                new Order { Val = (uint)seriesIndex },
                CreateSeriesTextLiteral(name),
                CreateXValuesFromRange(seriesRange.XRangeA1, defaultSheetName),
                CreateYValuesFromRange(seriesRange.YRangeA1, defaultSheetName)
            );
        }

        private static BubbleChart CreateBubbleChartFromRanges(IReadOnlyList<ExcelChartSeriesRange> seriesRanges,
            string defaultSheetName, uint xAxisId, uint yAxisId) {
            var bubbleChart = new BubbleChart(new VaryColors { Val = false });

            for (int i = 0; i < seriesRanges.Count; i++) {
                bubbleChart.Append(CreateBubbleChartSeriesFromRanges(i, seriesRanges[i], defaultSheetName));
            }

            bubbleChart.Append(CreateDefaultDataLabels());
            bubbleChart.Append(new BubbleScale { Val = 100 });
            bubbleChart.Append(new ShowNegativeBubbles { Val = false });
            bubbleChart.Append(new SizeRepresents { Val = SizeRepresentsValues.Area });
            bubbleChart.Append(new AxisId { Val = xAxisId });
            bubbleChart.Append(new AxisId { Val = yAxisId });
            return bubbleChart;
        }

        private static BubbleChartSeries CreateBubbleChartSeriesFromRanges(int seriesIndex, ExcelChartSeriesRange seriesRange,
            string defaultSheetName) {
            if (string.IsNullOrWhiteSpace(seriesRange.BubbleSizeRangeA1)) {
                throw new ArgumentException("Bubble size range is required for bubble charts.", nameof(seriesRange));
            }

            string name = string.IsNullOrWhiteSpace(seriesRange.Name) ? $"Series {seriesIndex + 1}" : seriesRange.Name;
            return new BubbleChartSeries(
                new ChartIndex { Val = (uint)seriesIndex },
                new Order { Val = (uint)seriesIndex },
                CreateSeriesTextLiteral(name),
                CreateXValuesFromRange(seriesRange.XRangeA1, defaultSheetName),
                CreateYValuesFromRange(seriesRange.YRangeA1, defaultSheetName),
                CreateBubbleSizeFromRange(seriesRange.BubbleSizeRangeA1!, defaultSheetName)
            );
        }

        private static XValues CreateXValuesFromRange(string rangeA1, string defaultSheetName) {
            string formula = EnsureSheetQualifiedRange(defaultSheetName, rangeA1);
            return new XValues(new NumberReference(new Formula { Text = formula }));
        }

        private static YValues CreateYValuesFromRange(string rangeA1, string defaultSheetName) {
            string formula = EnsureSheetQualifiedRange(defaultSheetName, rangeA1);
            return new YValues(new NumberReference(new Formula { Text = formula }));
        }

        private static BubbleSize CreateBubbleSizeFromRange(string rangeA1, string defaultSheetName) {
            string formula = EnsureSheetQualifiedRange(defaultSheetName, rangeA1);
            return new BubbleSize(new NumberReference(new Formula { Text = formula }));
        }

        private static CategoryAxis CreateCategoryAxis(uint axisId, uint crossingAxisId, AxisPositionValues? position = null) {
            AxisPositionValues axisPosition = position ?? AxisPositionValues.Bottom;
            return new CategoryAxis(
                new AxisId { Val = axisId },
                new Scaling(new Orientation { Val = OrientationValues.MinMax }),
                new Delete { Val = false },
                new AxisPosition { Val = axisPosition },
                new NumberingFormat { FormatCode = "General", SourceLinked = true },
                new MajorTickMark { Val = TickMarkValues.None },
                new MinorTickMark { Val = TickMarkValues.None },
                new TickLabelPosition { Val = TickLabelPositionValues.NextTo },
                new CrossingAxis { Val = crossingAxisId },
                new Crosses { Val = CrossesValues.AutoZero },
                new AutoLabeled { Val = true },
                new LabelAlignment { Val = LabelAlignmentValues.Center },
                new LabelOffset { Val = (UInt16Value)100U },
                new NoMultiLevelLabels { Val = false }
            );
        }

        private static ValueAxis CreateValueAxis(uint axisId, uint crossingAxisId, AxisPositionValues? position = null) {
            AxisPositionValues axisPosition = position ?? AxisPositionValues.Left;
            return new ValueAxis(
                new AxisId { Val = axisId },
                new Scaling(new Orientation { Val = OrientationValues.MinMax }),
                new Delete { Val = false },
                new AxisPosition { Val = axisPosition },
                new MajorGridlines(),
                new NumberingFormat { FormatCode = "General", SourceLinked = true },
                new MajorTickMark { Val = TickMarkValues.None },
                new MinorTickMark { Val = TickMarkValues.None },
                new TickLabelPosition { Val = TickLabelPositionValues.NextTo },
                new CrossingAxis { Val = crossingAxisId },
                new Crosses { Val = CrossesValues.AutoZero },
                new CrossBetween { Val = CrossBetweenValues.Between }
            );
        }

        private static DataLabels CreateDefaultDataLabels() {
            return new DataLabels(
                new ShowLegendKey { Val = false },
                new ShowValue { Val = false },
                new ShowCategoryName { Val = false },
                new ShowSeriesName { Val = false },
                new ShowPercent { Val = false },
                new ShowBubbleSize { Val = false }
            );
        }

        private static SeriesText CreateSeriesText(ExcelChartDataRange range, int seriesIndex, string name) {
            if (range.HasHeaderRow) {
                string seriesCell = range.SeriesNameCellA1(seriesIndex);
                string formula = BuildSheetQualifiedRange(range.SheetName, seriesCell);
                return new SeriesText(CreateStringReference(formula, new[] { name }));
            }

            return new SeriesText(new NumericValue { Text = name });
        }

        private static SeriesText CreateSeriesTextLiteral(string name) {
            return new SeriesText(new NumericValue { Text = name });
        }

        private static CategoryAxisData CreateCategoryAxisData(ExcelChartDataRange range) {
            string formula = BuildSheetQualifiedRange(range.SheetName, range.CategoriesRangeA1);
            return new CategoryAxisData(CreateStringReference(formula, null));
        }

        private static XValues CreateXValues(ExcelChartDataRange range, IReadOnlyList<double>? xValues) {
            string formula = BuildSheetQualifiedRange(range.SheetName, range.CategoriesRangeA1);
            return new XValues(CreateNumberReference(formula, xValues));
        }

        private static Values CreateValues(ExcelChartDataRange range, int seriesIndex, ExcelChartSeries? series) {
            string formula = BuildSheetQualifiedRange(range.SheetName, range.SeriesValuesRangeA1(seriesIndex));
            return new Values(CreateNumberReference(formula, series?.Values));
        }

        private static YValues CreateYValues(ExcelChartDataRange range, int seriesIndex, ExcelChartSeries? series) {
            string formula = BuildSheetQualifiedRange(range.SheetName, range.SeriesValuesRangeA1(seriesIndex));
            return new YValues(CreateNumberReference(formula, series?.Values));
        }
    }
}
