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
            ValidateSingleSeriesPieVariants(descriptors);
            SeriesDescriptorSummary summary = SummarizeSeriesDescriptors(descriptors);
            if (summary.HasBubble) {
                throw new NotSupportedException("Bubble charts require explicit X/Y/size ranges. Use AddBubbleChartFromRanges.");
            }
            if (summary.HasStock && (summary.HasMultipleTypes || summary.HasSecondary)) {
                throw new NotSupportedException("Stock charts cannot be combined with other chart types or secondary axes.");
            }
            if (summary.HasSurface && (summary.HasMultipleTypes || summary.HasSecondary)) {
                throw new NotSupportedException("Surface charts cannot be combined with other chart types or secondary axes.");
            }
            if (summary.HasLine3D && (summary.HasMultipleTypes || summary.HasSecondary)) {
                throw new NotSupportedException("3-D line charts cannot be combined with other chart types or secondary axes.");
            }
            if (summary.HasBar3D && (summary.HasMultipleTypes || summary.HasSecondary)) {
                throw new NotSupportedException("3-D bar and column charts cannot be combined with other chart types or secondary axes.");
            }
            if (summary.HasArea3D && (summary.HasMultipleTypes || summary.HasSecondary)) {
                throw new NotSupportedException("3-D area charts cannot be combined with other chart types or secondary axes.");
            }
            if (summary.HasScatter && summary.HasMultipleTypes) {
                throw new NotSupportedException("Scatter charts cannot be combined with other chart types.");
            }

            if (summary.HasScatter) {
                if (summary.HasMultipleTypes) {
                    throw new NotSupportedException("Scatter charts cannot be combined with other chart types.");
                }

                uint xAxisId = ExcelChartAxisIdGenerator.GetNextId();
                uint yAxisId = ExcelChartAxisIdGenerator.GetNextId();
                plotArea.Append(CreateScatterChart(range, descriptors, xAxisId, yAxisId, data));
                plotArea.Append(CreateValueAxis(xAxisId, yAxisId, AxisPositionValues.Bottom));
                plotArea.Append(CreateValueAxis(yAxisId, xAxisId, AxisPositionValues.Left));
            } else if (summary.HasMultipleTypes || summary.HasSecondary) {
                BuildComboPlotArea(plotArea, range, descriptors, summary);
            } else {
                ExcelChartType chartType = descriptors.Count > 0 ? descriptors[0].ChartType : type;
                uint categoryAxisId = ExcelChartAxisIdGenerator.GetNextId();
                uint valueAxisId = ExcelChartAxisIdGenerator.GetNextId();
                uint seriesAxisId = IsSurfaceChartType(chartType) || chartType == ExcelChartType.Line3D || IsBar3DChartType(chartType) || IsArea3DChartType(chartType) ? ExcelChartAxisIdGenerator.GetNextId() : 0;

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
                    case ExcelChartType.ColumnStacked100:
                        plotArea.Append(CreateBarChart(range, descriptors, BarDirectionValues.Column, BarGroupingValues.PercentStacked, categoryAxisId, valueAxisId));
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
                    case ExcelChartType.BarStacked100:
                        plotArea.Append(CreateBarChart(range, descriptors, BarDirectionValues.Bar, BarGroupingValues.PercentStacked, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId, AxisPositionValues.Left));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId, AxisPositionValues.Bottom));
                        break;
                    case ExcelChartType.Column3DClustered:
                        plotArea.Append(CreateBar3DChart(range, descriptors, BarDirectionValues.Column, BarGroupingValues.Clustered, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId));
                        plotArea.Append(CreateSeriesAxis(seriesAxisId, valueAxisId));
                        break;
                    case ExcelChartType.Column3DStacked:
                        plotArea.Append(CreateBar3DChart(range, descriptors, BarDirectionValues.Column, BarGroupingValues.Stacked, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId));
                        plotArea.Append(CreateSeriesAxis(seriesAxisId, valueAxisId));
                        break;
                    case ExcelChartType.Column3DStacked100:
                        plotArea.Append(CreateBar3DChart(range, descriptors, BarDirectionValues.Column, BarGroupingValues.PercentStacked, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId));
                        plotArea.Append(CreateSeriesAxis(seriesAxisId, valueAxisId));
                        break;
                    case ExcelChartType.Bar3DClustered:
                        plotArea.Append(CreateBar3DChart(range, descriptors, BarDirectionValues.Bar, BarGroupingValues.Clustered, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId, AxisPositionValues.Left));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId, AxisPositionValues.Bottom));
                        plotArea.Append(CreateSeriesAxis(seriesAxisId, valueAxisId));
                        break;
                    case ExcelChartType.Bar3DStacked:
                        plotArea.Append(CreateBar3DChart(range, descriptors, BarDirectionValues.Bar, BarGroupingValues.Stacked, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId, AxisPositionValues.Left));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId, AxisPositionValues.Bottom));
                        plotArea.Append(CreateSeriesAxis(seriesAxisId, valueAxisId));
                        break;
                    case ExcelChartType.Bar3DStacked100:
                        plotArea.Append(CreateBar3DChart(range, descriptors, BarDirectionValues.Bar, BarGroupingValues.PercentStacked, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId, AxisPositionValues.Left));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId, AxisPositionValues.Bottom));
                        plotArea.Append(CreateSeriesAxis(seriesAxisId, valueAxisId));
                        break;
                    case ExcelChartType.Line:
                        plotArea.Append(CreateLineChart(range, descriptors, GroupingValues.Standard, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId));
                        break;
                    case ExcelChartType.LineStacked:
                        plotArea.Append(CreateLineChart(range, descriptors, GroupingValues.Stacked, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId));
                        break;
                    case ExcelChartType.LineStacked100:
                        plotArea.Append(CreateLineChart(range, descriptors, GroupingValues.PercentStacked, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId));
                        break;
                    case ExcelChartType.Line3D:
                        plotArea.Append(CreateLine3DChart(range, descriptors, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId));
                        plotArea.Append(CreateSeriesAxis(seriesAxisId, valueAxisId));
                        break;
                    case ExcelChartType.Area:
                        plotArea.Append(CreateAreaChart(range, descriptors, GroupingValues.Standard, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId));
                        break;
                    case ExcelChartType.AreaStacked:
                        plotArea.Append(CreateAreaChart(range, descriptors, GroupingValues.Stacked, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId));
                        break;
                    case ExcelChartType.AreaStacked100:
                        plotArea.Append(CreateAreaChart(range, descriptors, GroupingValues.PercentStacked, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId));
                        break;
                    case ExcelChartType.Area3D:
                        plotArea.Append(CreateArea3DChart(range, descriptors, GroupingValues.Standard, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId));
                        plotArea.Append(CreateSeriesAxis(seriesAxisId, valueAxisId));
                        break;
                    case ExcelChartType.Area3DStacked:
                        plotArea.Append(CreateArea3DChart(range, descriptors, GroupingValues.Stacked, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId));
                        plotArea.Append(CreateSeriesAxis(seriesAxisId, valueAxisId));
                        break;
                    case ExcelChartType.Area3DStacked100:
                        plotArea.Append(CreateArea3DChart(range, descriptors, GroupingValues.PercentStacked, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId));
                        plotArea.Append(CreateSeriesAxis(seriesAxisId, valueAxisId));
                        break;
                    case ExcelChartType.Radar:
                        plotArea.Append(CreateRadarChart(range, descriptors, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId));
                        break;
                    case ExcelChartType.Stock:
                        plotArea.Append(CreateStockChart(range, descriptors, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId));
                        break;
                    case ExcelChartType.Surface:
                        plotArea.Append(CreateSurface3DChart(range, descriptors, false, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId));
                        plotArea.Append(CreateSeriesAxis(seriesAxisId, valueAxisId));
                        break;
                    case ExcelChartType.SurfaceWireframe:
                        plotArea.Append(CreateSurface3DChart(range, descriptors, true, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId));
                        plotArea.Append(CreateSeriesAxis(seriesAxisId, valueAxisId));
                        break;
                    case ExcelChartType.SurfaceContour:
                        plotArea.Append(CreateSurfaceChart(range, descriptors, false, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId));
                        plotArea.Append(CreateSeriesAxis(seriesAxisId, valueAxisId));
                        break;
                    case ExcelChartType.SurfaceContourWireframe:
                        plotArea.Append(CreateSurfaceChart(range, descriptors, true, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId));
                        plotArea.Append(CreateSeriesAxis(seriesAxisId, valueAxisId));
                        break;
                    case ExcelChartType.Pie:
                        plotArea.Append(CreatePieChart(range, descriptors));
                        break;
                    case ExcelChartType.Pie3D:
                        plotArea.Append(CreatePie3DChart(range, descriptors));
                        break;
                    case ExcelChartType.PieOfPie:
                        plotArea.Append(CreateOfPieChart(range, descriptors, OfPieValues.Pie));
                        break;
                    case ExcelChartType.BarOfPie:
                        plotArea.Append(CreateOfPieChart(range, descriptors, OfPieValues.Bar));
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
            if (type == ExcelChartType.Bubble && HasMissingBubbleSizeRange(seriesRanges)) {
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

        private static void BuildComboPlotArea(PlotArea plotArea, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors, SeriesDescriptorSummary summary) {
            if (summary.HasBubble) {
                throw new NotSupportedException("Bubble charts cannot be combined with other chart types.");
            }
            if (summary.HasStock) {
                throw new NotSupportedException("Stock charts cannot be combined with other chart types.");
            }
            if (summary.HasSurface) {
                throw new NotSupportedException("Surface charts cannot be combined with other chart types.");
            }
            if (summary.HasLine3D) {
                throw new NotSupportedException("3-D line charts cannot be combined with other chart types.");
            }
            if (summary.HasBar3D) {
                throw new NotSupportedException("3-D bar and column charts cannot be combined with other chart types.");
            }
            if (summary.HasArea3D) {
                throw new NotSupportedException("3-D area charts cannot be combined with other chart types.");
            }
            if (summary.HasScatter && summary.HasMultipleTypes) {
                throw new NotSupportedException("Scatter charts cannot be combined with other chart types.");
            }

            if (summary.HasBar && summary.HasNonBar) {
                throw new NotSupportedException("Cannot combine horizontal bar charts with other chart types.");
            }

            if (summary.HasPieOrDoughnut) {
                throw new NotSupportedException("Pie and doughnut charts cannot be combined with other chart types.");
            }

            bool isBarOrientation = summary.HasBar;
            AxisPositionValues primaryCategoryPosition = isBarOrientation ? AxisPositionValues.Left : AxisPositionValues.Bottom;
            AxisPositionValues primaryValuePosition = isBarOrientation ? AxisPositionValues.Bottom : AxisPositionValues.Left;
            AxisPositionValues secondaryCategoryPosition = isBarOrientation ? AxisPositionValues.Right : AxisPositionValues.Top;
            AxisPositionValues secondaryValuePosition = isBarOrientation ? AxisPositionValues.Top : AxisPositionValues.Right;

            uint primaryCategoryId = ExcelChartAxisIdGenerator.GetNextId();
            uint primaryValueId = ExcelChartAxisIdGenerator.GetNextId();
            uint secondaryCategoryId = summary.HasSecondary ? ExcelChartAxisIdGenerator.GetNextId() : 0;
            uint secondaryValueId = summary.HasSecondary ? ExcelChartAxisIdGenerator.GetNextId() : 0;

            foreach (SeriesDescriptorGroup group in GroupSeriesDescriptors(descriptors)) {
                uint categoryAxisId = group.AxisGroup == ExcelChartAxisGroup.Secondary ? secondaryCategoryId : primaryCategoryId;
                uint valueAxisId = group.AxisGroup == ExcelChartAxisGroup.Secondary ? secondaryValueId : primaryValueId;

                switch (group.ChartType) {
                    case ExcelChartType.ColumnClustered:
                    case ExcelChartType.ColumnStacked:
                    case ExcelChartType.ColumnStacked100:
                    case ExcelChartType.BarClustered:
                    case ExcelChartType.BarStacked:
                    case ExcelChartType.BarStacked100: {
                        var settings = GetBarChartSettings(group.ChartType);
                        plotArea.Append(CreateBarChart(range, group.Descriptors, settings.Direction, settings.Grouping, categoryAxisId, valueAxisId));
                        break;
                    }
                    case ExcelChartType.Line:
                    case ExcelChartType.LineStacked:
                    case ExcelChartType.LineStacked100:
                        plotArea.Append(CreateLineChart(range, group.Descriptors, GetLineGrouping(group.ChartType), categoryAxisId, valueAxisId));
                        break;
                    case ExcelChartType.Area:
                    case ExcelChartType.AreaStacked:
                    case ExcelChartType.AreaStacked100:
                        plotArea.Append(CreateAreaChart(range, group.Descriptors, GetAreaGrouping(group.ChartType), categoryAxisId, valueAxisId));
                        break;
                    default:
                        throw new NotSupportedException($"Chart type {group.ChartType} is not supported in combination charts.");
                }
            }

            plotArea.Append(CreateCategoryAxis(primaryCategoryId, primaryValueId, primaryCategoryPosition));
            plotArea.Append(CreateValueAxis(primaryValueId, primaryCategoryId, primaryValuePosition));
            if (summary.HasSecondary) {
                plotArea.Append(CreateCategoryAxis(secondaryCategoryId, secondaryValueId, secondaryCategoryPosition));
                plotArea.Append(CreateValueAxis(secondaryValueId, secondaryCategoryId, secondaryValuePosition));
            }
        }

        private static bool HasMissingBubbleSizeRange(IReadOnlyList<ExcelChartSeriesRange> seriesRanges) {
            for (int i = 0; i < seriesRanges.Count; i++) {
                if (string.IsNullOrWhiteSpace(seriesRanges[i].BubbleSizeRangeA1)) {
                    return true;
                }
            }

            return false;
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

        private static Bar3DChart CreateBar3DChart(ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors,
            BarDirectionValues direction, BarGroupingValues grouping, uint categoryAxisId, uint valueAxisId, uint seriesAxisId) {
            var barChart = new Bar3DChart(
                new BarDirection { Val = direction },
                new BarGrouping { Val = grouping },
                new VaryColors { Val = false });

            foreach (var descriptor in seriesDescriptors) {
                barChart.Append(CreateBarChartSeries(descriptor.Index, range, descriptor.Series));
            }

            barChart.Append(CreateDefaultDataLabels());
            barChart.Append(new GapWidth { Val = (UInt16Value)150U });
            barChart.Append(new GapDepth { Val = (UInt16Value)150U });
            barChart.Append(new Shape { Val = ShapeValues.Box });
            barChart.Append(new AxisId { Val = categoryAxisId });
            barChart.Append(new AxisId { Val = valueAxisId });
            barChart.Append(new AxisId { Val = seriesAxisId });
            return barChart;
        }

        private static LineChart CreateLineChart(ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors,
            GroupingValues grouping, uint categoryAxisId, uint valueAxisId) {
            var lineChart = new LineChart(
                new Grouping { Val = grouping },
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

        private static Line3DChart CreateLine3DChart(ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors,
            uint categoryAxisId, uint valueAxisId, uint seriesAxisId) {
            var lineChart = new Line3DChart(
                new Grouping { Val = GroupingValues.Standard },
                new VaryColors { Val = false });

            foreach (var descriptor in seriesDescriptors) {
                lineChart.Append(CreateLineChartSeries(descriptor.Index, range, descriptor.Series));
            }

            lineChart.Append(CreateDefaultDataLabels());
            lineChart.Append(new GapDepth { Val = (UInt16Value)150U });
            lineChart.Append(new AxisId { Val = categoryAxisId });
            lineChart.Append(new AxisId { Val = valueAxisId });
            lineChart.Append(new AxisId { Val = seriesAxisId });
            return lineChart;
        }

        private static AreaChart CreateAreaChart(ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors,
            GroupingValues grouping, uint categoryAxisId, uint valueAxisId) {
            var areaChart = new AreaChart(
                new Grouping { Val = grouping },
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

        private static Area3DChart CreateArea3DChart(ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors,
            GroupingValues grouping, uint categoryAxisId, uint valueAxisId, uint seriesAxisId) {
            var areaChart = new Area3DChart(
                new Grouping { Val = grouping },
                new VaryColors { Val = false });

            foreach (var descriptor in seriesDescriptors) {
                areaChart.Append(CreateAreaChartSeries(descriptor.Index, range, descriptor.Series));
            }

            areaChart.Append(CreateDefaultDataLabels());
            areaChart.Append(new GapDepth { Val = (UInt16Value)150U });
            areaChart.Append(new AxisId { Val = categoryAxisId });
            areaChart.Append(new AxisId { Val = valueAxisId });
            areaChart.Append(new AxisId { Val = seriesAxisId });
            return areaChart;
        }

        private static RadarChart CreateRadarChart(ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors,
            uint categoryAxisId, uint valueAxisId) {
            var radarChart = new RadarChart(
                new RadarStyle { Val = RadarStyleValues.Standard },
                new VaryColors { Val = false });

            foreach (var descriptor in seriesDescriptors) {
                radarChart.Append(CreateRadarChartSeries(descriptor.Index, range, descriptor.Series));
            }

            radarChart.Append(CreateDefaultDataLabels());
            radarChart.Append(new AxisId { Val = categoryAxisId });
            radarChart.Append(new AxisId { Val = valueAxisId });
            return radarChart;
        }

        private static RadarChartSeries CreateRadarChartSeries(int seriesIndex, ExcelChartDataRange range, ExcelChartSeries? series) {
            return new RadarChartSeries(
                new ChartIndex { Val = (uint)seriesIndex },
                new Order { Val = (uint)seriesIndex },
                CreateSeriesText(range, seriesIndex, series?.Name ?? $"Series {seriesIndex + 1}"),
                CreateCategoryAxisData(range),
                CreateValues(range, seriesIndex, series)
            );
        }

        private static StockChart CreateStockChart(ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors,
            uint categoryAxisId, uint valueAxisId) {
            if (seriesDescriptors.Count < 3 || seriesDescriptors.Count > 4) {
                throw new ArgumentException("Stock charts require three series (high-low-close) or four series (open-high-low-close).", nameof(seriesDescriptors));
            }

            var stockChart = new StockChart();
            foreach (var descriptor in seriesDescriptors) {
                stockChart.Append(CreateLineChartSeries(descriptor.Index, range, descriptor.Series));
            }

            stockChart.Append(CreateDefaultDataLabels());
            stockChart.Append(new HighLowLines());
            if (seriesDescriptors.Count == 4) {
                stockChart.Append(new UpDownBars(
                    new GapWidth { Val = (UInt16Value)150U },
                    new UpBars(),
                    new DownBars()));
            }
            stockChart.Append(new AxisId { Val = categoryAxisId });
            stockChart.Append(new AxisId { Val = valueAxisId });
            return stockChart;
        }

        private static Surface3DChart CreateSurface3DChart(ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors,
            bool wireframe, uint categoryAxisId, uint valueAxisId, uint seriesAxisId) {
            var surfaceChart = new Surface3DChart(
                new Wireframe { Val = wireframe },
                new VaryColors { Val = false });

            foreach (var descriptor in seriesDescriptors) {
                surfaceChart.Append(CreateSurfaceChartSeries(descriptor.Index, range, descriptor.Series));
            }

            surfaceChart.Append(new AxisId { Val = categoryAxisId });
            surfaceChart.Append(new AxisId { Val = valueAxisId });
            surfaceChart.Append(new AxisId { Val = seriesAxisId });
            return surfaceChart;
        }

        private static SurfaceChart CreateSurfaceChart(ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors,
            bool wireframe, uint categoryAxisId, uint valueAxisId, uint seriesAxisId) {
            var surfaceChart = new SurfaceChart(
                new Wireframe { Val = wireframe });

            foreach (var descriptor in seriesDescriptors) {
                surfaceChart.Append(CreateSurfaceChartSeries(descriptor.Index, range, descriptor.Series));
            }

            surfaceChart.Append(new AxisId { Val = categoryAxisId });
            surfaceChart.Append(new AxisId { Val = valueAxisId });
            surfaceChart.Append(new AxisId { Val = seriesAxisId });
            return surfaceChart;
        }

        private static SurfaceChartSeries CreateSurfaceChartSeries(int seriesIndex, ExcelChartDataRange range, ExcelChartSeries? series) {
            return new SurfaceChartSeries(
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

        private static Pie3DChart CreatePie3DChart(ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors) {
            var pieChart = new Pie3DChart(new VaryColors { Val = true });

            foreach (var descriptor in seriesDescriptors) {
                pieChart.Append(CreatePieChartSeries(descriptor.Index, range, descriptor.Series));
            }

            pieChart.Append(CreateDefaultDataLabels());
            return pieChart;
        }

        private static OfPieChart CreateOfPieChart(ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors, OfPieValues type) {
            var chart = new OfPieChart(
                new OfPieType { Val = type },
                new VaryColors { Val = true });

            foreach (var descriptor in seriesDescriptors) {
                chart.Append(CreatePieChartSeries(descriptor.Index, range, descriptor.Series));
            }

            chart.Append(CreateDefaultDataLabels());
            chart.Append(new GapWidth { Val = (UInt16Value)150U });
            chart.Append(new SplitType { Val = SplitValues.Position });
            chart.Append(new SplitPosition { Val = 2D });
            chart.Append(new SecondPieSize { Val = (UInt16Value)75U });
            chart.Append(new SeriesLines());
            return chart;
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
            var chart = new DoughnutChart(new VaryColors { Val = true });

            foreach (var descriptor in seriesDescriptors) {
                chart.Append(CreatePieChartSeries(descriptor.Index, range, descriptor.Series));
            }

            chart.Append(CreateDefaultDataLabels());
            chart.Append(new HoleSize { Val = 50 });
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

        private static SeriesAxis CreateSeriesAxis(uint axisId, uint crossingAxisId, AxisPositionValues? position = null) {
            AxisPositionValues axisPosition = position ?? AxisPositionValues.Right;
            return new SeriesAxis(
                new AxisId { Val = axisId },
                new Scaling(new Orientation { Val = OrientationValues.MinMax }),
                new Delete { Val = false },
                new AxisPosition { Val = axisPosition },
                new MajorTickMark { Val = TickMarkValues.None },
                new MinorTickMark { Val = TickMarkValues.None },
                new TickLabelPosition { Val = TickLabelPositionValues.NextTo },
                new CrossingAxis { Val = crossingAxisId }
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
                return new SeriesText(CreateSingleStringReference(formula, name));
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
