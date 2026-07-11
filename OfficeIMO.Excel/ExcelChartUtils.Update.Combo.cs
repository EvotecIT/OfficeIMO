using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using ChartIndex = DocumentFormat.OpenXml.Drawing.Charts.Index;

namespace OfficeIMO.Excel {
    internal static partial class ExcelChartUtils {
        private static void UpdateComboChartData(PlotArea plotArea, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            if (plotArea == null) {
                throw new ArgumentNullException(nameof(plotArea));
            }

            SeriesDescriptorSummary summary = SummarizeSeriesDescriptors(descriptors);
            if (summary.HasStock) {
                throw new NotSupportedException("Stock charts cannot be used in combination chart updates.");
            }
            if (summary.HasSurface) {
                throw new NotSupportedException("Surface charts cannot be used in combination chart updates.");
            }
            if (summary.HasLine3D) {
                throw new NotSupportedException("3-D line charts cannot be used in combination chart updates.");
            }
            if (summary.HasBar3D) {
                throw new NotSupportedException("3-D bar and column charts cannot be used in combination chart updates.");
            }
            if (summary.HasArea3D) {
                throw new NotSupportedException("3-D area charts cannot be used in combination chart updates.");
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

            var axisIds = EnsureAxisPairs(plotArea, summary.HasSecondary, primaryCategoryPosition, primaryValuePosition, secondaryCategoryPosition, secondaryValuePosition);

            var usedCharts = new List<OpenXmlCompositeElement>();
            foreach (SeriesDescriptorGroup group in GroupSeriesDescriptors(descriptors)) {
                uint categoryAxisId = group.AxisGroup == ExcelChartAxisGroup.Secondary ? axisIds.SecondaryCategoryId : axisIds.PrimaryCategoryId;
                uint valueAxisId = group.AxisGroup == ExcelChartAxisGroup.Secondary ? axisIds.SecondaryValueId : axisIds.PrimaryValueId;

                switch (group.ChartType) {
                    case ExcelChartType.ColumnClustered:
                        usedCharts.Add(UpdateOrCreateBarChart(plotArea, group.Descriptors, data, range, BarDirectionValues.Column, BarGroupingValues.Clustered, categoryAxisId, valueAxisId));
                        break;
                    case ExcelChartType.ColumnStacked:
                        usedCharts.Add(UpdateOrCreateBarChart(plotArea, group.Descriptors, data, range, BarDirectionValues.Column, BarGroupingValues.Stacked, categoryAxisId, valueAxisId));
                        break;
                    case ExcelChartType.ColumnStacked100:
                        usedCharts.Add(UpdateOrCreateBarChart(plotArea, group.Descriptors, data, range, BarDirectionValues.Column, BarGroupingValues.PercentStacked, categoryAxisId, valueAxisId));
                        break;
                    case ExcelChartType.BarClustered:
                        usedCharts.Add(UpdateOrCreateBarChart(plotArea, group.Descriptors, data, range, BarDirectionValues.Bar, BarGroupingValues.Clustered, categoryAxisId, valueAxisId));
                        break;
                    case ExcelChartType.BarStacked:
                        usedCharts.Add(UpdateOrCreateBarChart(plotArea, group.Descriptors, data, range, BarDirectionValues.Bar, BarGroupingValues.Stacked, categoryAxisId, valueAxisId));
                        break;
                    case ExcelChartType.BarStacked100:
                        usedCharts.Add(UpdateOrCreateBarChart(plotArea, group.Descriptors, data, range, BarDirectionValues.Bar, BarGroupingValues.PercentStacked, categoryAxisId, valueAxisId));
                        break;
                    case ExcelChartType.Line:
                    case ExcelChartType.LineStacked:
                    case ExcelChartType.LineStacked100:
                        usedCharts.Add(UpdateOrCreateLineChart(plotArea, group.Descriptors, data, range, GetLineGrouping(group.ChartType), categoryAxisId, valueAxisId));
                        break;
                    case ExcelChartType.Area:
                    case ExcelChartType.AreaStacked:
                    case ExcelChartType.AreaStacked100:
                        usedCharts.Add(UpdateOrCreateAreaChart(plotArea, group.Descriptors, data, range, GetAreaGrouping(group.ChartType), categoryAxisId, valueAxisId));
                        break;
                    case ExcelChartType.Scatter:
                        usedCharts.Add(UpdateOrCreateScatterChart(plotArea, group.Descriptors, data, range, categoryAxisId, valueAxisId));
                        break;
                    default:
                        throw new NotSupportedException($"Chart type {group.ChartType} is not supported in combination charts.");
                }
            }

            foreach (var chart in plotArea.Elements<BarChart>().Cast<OpenXmlCompositeElement>().ToList()) {
                if (!usedCharts.Contains(chart)) chart.Remove();
            }
            foreach (var chart in plotArea.Elements<LineChart>().Cast<OpenXmlCompositeElement>().ToList()) {
                if (!usedCharts.Contains(chart)) chart.Remove();
            }
            foreach (var chart in plotArea.Elements<AreaChart>().Cast<OpenXmlCompositeElement>().ToList()) {
                if (!usedCharts.Contains(chart)) chart.Remove();
            }
            foreach (var chart in plotArea.Elements<ScatterChart>().Cast<OpenXmlCompositeElement>().ToList()) {
                if (!usedCharts.Contains(chart)) chart.Remove();
            }
        }

        private static AxisIdSet EnsureAxisPairs(PlotArea plotArea, bool includeSecondary, AxisPositionValues primaryCategoryPosition,
            AxisPositionValues primaryValuePosition, AxisPositionValues secondaryCategoryPosition, AxisPositionValues secondaryValuePosition) {
            var categoryAxes = plotArea.Elements<CategoryAxis>().ToList();
            var valueAxes = plotArea.Elements<ValueAxis>().ToList();

            CategoryAxis? primaryCategory = categoryAxes.FirstOrDefault(ax => ax.AxisPosition?.Val?.Value == primaryCategoryPosition)
                ?? categoryAxes.FirstOrDefault();
            ValueAxis? primaryValue = valueAxes.FirstOrDefault(ax => ax.AxisPosition?.Val?.Value == primaryValuePosition)
                ?? valueAxes.FirstOrDefault();

            uint primaryCategoryId = primaryCategory?.AxisId?.Val?.Value ?? ExcelChartAxisIdGenerator.GetNextId();
            uint primaryValueId = primaryValue?.AxisId?.Val?.Value ?? ExcelChartAxisIdGenerator.GetNextId();

            if (primaryCategory == null) {
                primaryCategory = CreateCategoryAxis(primaryCategoryId, primaryValueId, primaryCategoryPosition);
                plotArea.Append(primaryCategory);
            } else {
                EnsureAxisId(primaryCategory, primaryCategoryId);
                EnsureAxisPosition(primaryCategory, primaryCategoryPosition);
                EnsureCrossingAxis(primaryCategory, primaryValueId);
            }

            if (primaryValue == null) {
                primaryValue = CreateValueAxis(primaryValueId, primaryCategoryId, primaryValuePosition);
                plotArea.Append(primaryValue);
            } else {
                EnsureAxisId(primaryValue, primaryValueId);
                EnsureAxisPosition(primaryValue, primaryValuePosition);
                EnsureCrossingAxis(primaryValue, primaryCategoryId);
            }

            uint secondaryCategoryId = 0;
            uint secondaryValueId = 0;
            if (includeSecondary) {
                CategoryAxis? secondaryCategory = categoryAxes.FirstOrDefault(ax => ax.AxisPosition?.Val?.Value == secondaryCategoryPosition);
                ValueAxis? secondaryValue = valueAxes.FirstOrDefault(ax => ax.AxisPosition?.Val?.Value == secondaryValuePosition);

                secondaryCategoryId = secondaryCategory?.AxisId?.Val?.Value ?? ExcelChartAxisIdGenerator.GetNextId();
                secondaryValueId = secondaryValue?.AxisId?.Val?.Value ?? ExcelChartAxisIdGenerator.GetNextId();

                if (secondaryCategory == null) {
                    secondaryCategory = CreateCategoryAxis(secondaryCategoryId, secondaryValueId, secondaryCategoryPosition);
                    plotArea.Append(secondaryCategory);
                } else {
                    EnsureAxisId(secondaryCategory, secondaryCategoryId);
                    EnsureAxisPosition(secondaryCategory, secondaryCategoryPosition);
                    EnsureCrossingAxis(secondaryCategory, secondaryValueId);
                }

                if (secondaryValue == null) {
                    secondaryValue = CreateValueAxis(secondaryValueId, secondaryCategoryId, secondaryValuePosition);
                    plotArea.Append(secondaryValue);
                } else {
                    EnsureAxisId(secondaryValue, secondaryValueId);
                    EnsureAxisPosition(secondaryValue, secondaryValuePosition);
                    EnsureCrossingAxis(secondaryValue, secondaryCategoryId);
                }
            }

            return new AxisIdSet(primaryCategoryId, primaryValueId, secondaryCategoryId, secondaryValueId);
        }

        private static void EnsureAxisId(OpenXmlCompositeElement axis, uint axisId) {
            AxisId id = axis.GetFirstChild<AxisId>() ?? new AxisId();
            id.Val = axisId;
            if (id.Parent == null) {
                axis.PrependChild(id);
            }
        }

        private static void EnsureAxisPosition(OpenXmlCompositeElement axis, AxisPositionValues position) {
            AxisPosition axisPosition = axis.GetFirstChild<AxisPosition>() ?? new AxisPosition();
            axisPosition.Val = position;
            if (axisPosition.Parent == null) {
                axis.Append(axisPosition);
            }
        }

        private static void EnsureCrossingAxis(OpenXmlCompositeElement axis, uint crossAxisId) {
            CrossingAxis crossingAxis = axis.GetFirstChild<CrossingAxis>() ?? new CrossingAxis();
            crossingAxis.Val = crossAxisId;
            if (crossingAxis.Parent == null) {
                axis.Append(crossingAxis);
            }
        }

        private static BarChart UpdateOrCreateBarChart(PlotArea plotArea, IReadOnlyList<SeriesDescriptor> descriptors, ExcelChartData data, ExcelChartDataRange range,
            BarDirectionValues direction, BarGroupingValues grouping, uint categoryAxisId, uint valueAxisId) {
            BarChart? chart = plotArea.Elements<BarChart>()
                .FirstOrDefault(c => (c.GetFirstChild<BarDirection>()?.Val ?? BarDirectionValues.Column) == direction
                                     && (c.GetFirstChild<BarGrouping>()?.Val ?? BarGroupingValues.Clustered) == grouping
                                     && ChartHasAxisIds(c, categoryAxisId, valueAxisId));

            chart ??= plotArea.Elements<BarChart>()
                .FirstOrDefault(c => (c.GetFirstChild<BarDirection>()?.Val ?? BarDirectionValues.Column) == direction
                                     && (c.GetFirstChild<BarGrouping>()?.Val ?? BarGroupingValues.Clustered) == grouping);

            if (chart == null) {
                chart = CreateBarChart(range, descriptors, direction, grouping, categoryAxisId, valueAxisId);
                plotArea.Append(chart);
            } else {
                ResetAxisIds(chart, categoryAxisId, valueAxisId);
                UpdateBarChartSeries(chart, data, range, descriptors);
            }

            return chart;
        }

        private static LineChart UpdateOrCreateLineChart(PlotArea plotArea, IReadOnlyList<SeriesDescriptor> descriptors, ExcelChartData data, ExcelChartDataRange range,
            GroupingValues grouping, uint categoryAxisId, uint valueAxisId) {
            LineChart? chart = plotArea.Elements<LineChart>()
                .FirstOrDefault(c => (c.GetFirstChild<Grouping>()?.Val ?? GroupingValues.Standard) == grouping
                                     && ChartHasAxisIds(c, categoryAxisId, valueAxisId));

            chart ??= plotArea.Elements<LineChart>()
                .FirstOrDefault(c => (c.GetFirstChild<Grouping>()?.Val ?? GroupingValues.Standard) == grouping);

            if (chart == null) {
                chart = CreateLineChart(range, descriptors, grouping, categoryAxisId, valueAxisId);
                plotArea.Append(chart);
            } else {
                EnsureGrouping(chart, grouping);
                ResetAxisIds(chart, categoryAxisId, valueAxisId);
                UpdateLineChartSeries(chart, data, range, descriptors);
            }

            return chart;
        }

        private static AreaChart UpdateOrCreateAreaChart(PlotArea plotArea, IReadOnlyList<SeriesDescriptor> descriptors, ExcelChartData data, ExcelChartDataRange range,
            GroupingValues grouping, uint categoryAxisId, uint valueAxisId) {
            AreaChart? chart = plotArea.Elements<AreaChart>()
                .FirstOrDefault(c => (c.GetFirstChild<Grouping>()?.Val ?? GroupingValues.Standard) == grouping
                                     && ChartHasAxisIds(c, categoryAxisId, valueAxisId));

            chart ??= plotArea.Elements<AreaChart>()
                .FirstOrDefault(c => (c.GetFirstChild<Grouping>()?.Val ?? GroupingValues.Standard) == grouping);

            if (chart == null) {
                chart = CreateAreaChart(range, descriptors, grouping, categoryAxisId, valueAxisId);
                plotArea.Append(chart);
            } else {
                EnsureGrouping(chart, grouping);
                ResetAxisIds(chart, categoryAxisId, valueAxisId);
                UpdateAreaChartSeries(chart, data, range, descriptors);
            }

            return chart;
        }

        private static void EnsureGrouping(OpenXmlCompositeElement chart, GroupingValues groupingValue) {
            Grouping grouping = chart.GetFirstChild<Grouping>() ?? new Grouping();
            grouping.Val = groupingValue;
            if (grouping.Parent == null) {
                chart.PrependChild(grouping);
            }
        }

        private static ScatterChart UpdateOrCreateScatterChart(PlotArea plotArea, IReadOnlyList<SeriesDescriptor> descriptors, ExcelChartData data, ExcelChartDataRange range,
            uint xAxisId, uint yAxisId) {
            ScatterChart? chart = plotArea.Elements<ScatterChart>()
                .FirstOrDefault(c => ChartHasAxisIds(c, xAxisId, yAxisId))
                ?? plotArea.Elements<ScatterChart>().FirstOrDefault();

            if (chart == null) {
                chart = CreateScatterChart(range, descriptors, xAxisId, yAxisId, data);
                plotArea.Append(chart);
            } else {
                ResetAxisIds(chart, xAxisId, yAxisId);
                UpdateScatterChartSeries(chart, data, range, descriptors);
            }

            return chart;
        }

        private static bool ChartHasAxisIds(OpenXmlCompositeElement chart, uint categoryAxisId, uint valueAxisId) {
            bool hasCategoryAxis = false;
            bool hasValueAxis = false;
            foreach (AxisId id in chart.Elements<AxisId>()) {
                uint? value = id.Val?.Value;
                if (value == categoryAxisId) {
                    hasCategoryAxis = true;
                    if (hasValueAxis) {
                        return true;
                    }
                } else if (value == valueAxisId) {
                    hasValueAxis = true;
                    if (hasCategoryAxis) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static void ResetAxisIds(OpenXmlCompositeElement chart, uint categoryAxisId, uint valueAxisId) {
            chart.RemoveAllChildren<AxisId>();
            chart.Append(new AxisId { Val = categoryAxisId });
            chart.Append(new AxisId { Val = valueAxisId });
        }

    }
}
