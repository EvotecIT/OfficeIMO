using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    internal static partial class ExcelChartUtils {
        private static void UpdateComboChartData(OfficeXmlNode plotArea, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            if (plotArea.IsNone) throw new ArgumentNullException(nameof(plotArea));

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
            string primaryCategoryPosition = isBarOrientation ? AxisPositionLeft : AxisPositionBottom;
            string primaryValuePosition = isBarOrientation ? AxisPositionBottom : AxisPositionLeft;
            string secondaryCategoryPosition = isBarOrientation ? AxisPositionRight : AxisPositionTop;
            string secondaryValuePosition = isBarOrientation ? AxisPositionTop : AxisPositionRight;

            var axisIds = EnsureAxisPairs(plotArea, summary.HasSecondary, primaryCategoryPosition, primaryValuePosition, secondaryCategoryPosition, secondaryValuePosition);

            var usedCharts = new List<OfficeXmlNode>();
            foreach (SeriesDescriptorGroup group in GroupSeriesDescriptors(descriptors)) {
                uint categoryAxisId = group.AxisGroup == OfficeChartAxisGroup.Secondary ? axisIds.SecondaryCategoryId : axisIds.PrimaryCategoryId;
                uint valueAxisId = group.AxisGroup == OfficeChartAxisGroup.Secondary ? axisIds.SecondaryValueId : axisIds.PrimaryValueId;

                switch (group.ChartType) {
                    case OfficeChartType.ColumnClustered:
                        usedCharts.Add(UpdateOrCreateBarChart(plotArea, group.Descriptors, data, range, BarDirectionColumn, BarGroupingClustered, categoryAxisId, valueAxisId));
                        break;
                    case OfficeChartType.ColumnStacked:
                        usedCharts.Add(UpdateOrCreateBarChart(plotArea, group.Descriptors, data, range, BarDirectionColumn, GroupingStacked, categoryAxisId, valueAxisId));
                        break;
                    case OfficeChartType.ColumnStacked100:
                        usedCharts.Add(UpdateOrCreateBarChart(plotArea, group.Descriptors, data, range, BarDirectionColumn, GroupingPercentStacked, categoryAxisId, valueAxisId));
                        break;
                    case OfficeChartType.BarClustered:
                        usedCharts.Add(UpdateOrCreateBarChart(plotArea, group.Descriptors, data, range, BarDirectionBar, BarGroupingClustered, categoryAxisId, valueAxisId));
                        break;
                    case OfficeChartType.BarStacked:
                        usedCharts.Add(UpdateOrCreateBarChart(plotArea, group.Descriptors, data, range, BarDirectionBar, GroupingStacked, categoryAxisId, valueAxisId));
                        break;
                    case OfficeChartType.BarStacked100:
                        usedCharts.Add(UpdateOrCreateBarChart(plotArea, group.Descriptors, data, range, BarDirectionBar, GroupingPercentStacked, categoryAxisId, valueAxisId));
                        break;
                    case OfficeChartType.Line:
                    case OfficeChartType.LineStacked:
                    case OfficeChartType.LineStacked100:
                        usedCharts.Add(UpdateOrCreateLineChart(plotArea, group.Descriptors, data, range, GetLineGrouping(group.ChartType), categoryAxisId, valueAxisId));
                        break;
                    case OfficeChartType.Area:
                    case OfficeChartType.AreaStacked:
                    case OfficeChartType.AreaStacked100:
                        usedCharts.Add(UpdateOrCreateAreaChart(plotArea, group.Descriptors, data, range, GetAreaGrouping(group.ChartType), categoryAxisId, valueAxisId));
                        break;
                    case OfficeChartType.Scatter:
                        usedCharts.Add(UpdateOrCreateScatterChart(plotArea, group.Descriptors, data, range, categoryAxisId, valueAxisId));
                        break;
                    default:
                        throw new NotSupportedException($"Chart type {group.ChartType} is not supported in combination charts.");
                }
            }

            RemoveUnusedCharts(plotArea, "barChart", usedCharts);
            RemoveUnusedCharts(plotArea, "lineChart", usedCharts);
            RemoveUnusedCharts(plotArea, "areaChart", usedCharts);
            RemoveUnusedCharts(plotArea, "scatterChart", usedCharts);
        }

        private static void RemoveUnusedCharts(OfficeXmlNode plotArea, string localName, List<OfficeXmlNode> usedCharts) {
            foreach (OfficeXmlNode chart in plotArea.Children(C, localName).ToList()) {
                if (!usedCharts.Contains(chart)) chart.Remove();
            }
        }

        private static AxisIdSet EnsureAxisPairs(OfficeXmlNode plotArea, bool includeSecondary, string primaryCategoryPosition,
            string primaryValuePosition, string secondaryCategoryPosition, string secondaryValuePosition) {
            List<OfficeXmlNode> categoryAxes = plotArea.Children(C, "catAx").ToList();
            List<OfficeXmlNode> valueAxes = plotArea.Children(C, "valAx").ToList();

            OfficeXmlNode primaryCategory = FindAxis(categoryAxes, primaryCategoryPosition, fallBackToFirst: true);
            OfficeXmlNode primaryValue = FindAxis(valueAxes, primaryValuePosition, fallBackToFirst: true);

            uint primaryCategoryId = primaryCategory.Child(C, "axId").UInt32Val() ?? ExcelChartAxisIdGenerator.GetNextId();
            uint primaryValueId = primaryValue.Child(C, "axId").UInt32Val() ?? ExcelChartAxisIdGenerator.GetNextId();

            if (primaryCategory.IsNone) {
                plotArea.Append(CreateCategoryAxis(plotArea, primaryCategoryId, primaryValueId, primaryCategoryPosition));
            } else {
                EnsureAxisId(primaryCategory, primaryCategoryId);
                EnsureAxisPosition(primaryCategory, primaryCategoryPosition);
                EnsureCrossingAxis(primaryCategory, primaryValueId);
            }

            if (primaryValue.IsNone) {
                plotArea.Append(CreateValueAxis(plotArea, primaryValueId, primaryCategoryId, primaryValuePosition));
            } else {
                EnsureAxisId(primaryValue, primaryValueId);
                EnsureAxisPosition(primaryValue, primaryValuePosition);
                EnsureCrossingAxis(primaryValue, primaryCategoryId);
            }

            uint secondaryCategoryId = 0;
            uint secondaryValueId = 0;
            if (includeSecondary) {
                OfficeXmlNode secondaryCategory = FindAxis(categoryAxes, secondaryCategoryPosition, fallBackToFirst: false);
                OfficeXmlNode secondaryValue = FindAxis(valueAxes, secondaryValuePosition, fallBackToFirst: false);

                secondaryCategoryId = secondaryCategory.Child(C, "axId").UInt32Val() ?? ExcelChartAxisIdGenerator.GetNextId();
                secondaryValueId = secondaryValue.Child(C, "axId").UInt32Val() ?? ExcelChartAxisIdGenerator.GetNextId();

                if (secondaryCategory.IsNone) {
                    plotArea.Append(CreateCategoryAxis(plotArea, secondaryCategoryId, secondaryValueId, secondaryCategoryPosition));
                } else {
                    EnsureAxisId(secondaryCategory, secondaryCategoryId);
                    EnsureAxisPosition(secondaryCategory, secondaryCategoryPosition);
                    EnsureCrossingAxis(secondaryCategory, secondaryValueId);
                }

                if (secondaryValue.IsNone) {
                    plotArea.Append(CreateValueAxis(plotArea, secondaryValueId, secondaryCategoryId, secondaryValuePosition));
                } else {
                    EnsureAxisId(secondaryValue, secondaryValueId);
                    EnsureAxisPosition(secondaryValue, secondaryValuePosition);
                    EnsureCrossingAxis(secondaryValue, secondaryCategoryId);
                }
            }

            return new AxisIdSet(primaryCategoryId, primaryValueId, secondaryCategoryId, secondaryValueId);
        }

        // The first axis at this position, or the first axis at all when asked to fall back to it.
        private static OfficeXmlNode FindAxis(List<OfficeXmlNode> axes, string position, bool fallBackToFirst) {
            foreach (OfficeXmlNode axis in axes) {
                if (axis.ChildVal(C, "axPos") == position) return axis;
            }
            return fallBackToFirst && axes.Count > 0 ? axes[0] : OfficeXmlNode.None;
        }

        private static void EnsureAxisId(OfficeXmlNode axis, uint axisId) {
            OfficeXmlNode id = axis.Child(C, "axId");
            if (id.IsNone) id = axis.InsertBefore(axis.New(C, "axId"), axis.FirstChild);
            id.SetVal(axisId);
        }

        private static void EnsureAxisPosition(OfficeXmlNode axis, string position) {
            OfficeXmlNode axisPosition = axis.Child(C, "axPos");
            if (axisPosition.IsNone) axisPosition = axis.Append(axis.New(C, "axPos"));
            axisPosition.SetVal(position);
        }

        private static void EnsureCrossingAxis(OfficeXmlNode axis, uint crossAxisId) {
            OfficeXmlNode crossingAxis = axis.Child(C, "crossAx");
            if (crossingAxis.IsNone) crossingAxis = axis.Append(axis.New(C, "crossAx"));
            crossingAxis.SetVal(crossAxisId);
        }

        private static bool IsBarChartOf(OfficeXmlNode chart, string direction, string grouping) =>
            (chart.ChildVal(C, "barDir") ?? BarDirectionColumn) == direction
            && (chart.ChildVal(C, "grouping") ?? BarGroupingClustered) == grouping;

        private static bool IsGroupedChartOf(OfficeXmlNode chart, string grouping) =>
            (chart.ChildVal(C, "grouping") ?? GroupingStandard) == grouping;

        private static OfficeXmlNode UpdateOrCreateBarChart(OfficeXmlNode plotArea, IReadOnlyList<SeriesDescriptor> descriptors, ExcelChartData data, ExcelChartDataRange range,
            string direction, string grouping, uint categoryAxisId, uint valueAxisId) {
            OfficeXmlNode chart = OfficeXmlNode.None;
            foreach (OfficeXmlNode candidate in plotArea.Children(C, "barChart")) {
                if (IsBarChartOf(candidate, direction, grouping) && ChartHasAxisIds(candidate, categoryAxisId, valueAxisId)) {
                    chart = candidate;
                    break;
                }
            }
            if (chart.IsNone) {
                foreach (OfficeXmlNode candidate in plotArea.Children(C, "barChart")) {
                    if (IsBarChartOf(candidate, direction, grouping)) {
                        chart = candidate;
                        break;
                    }
                }
            }

            if (chart.IsNone) {
                chart = plotArea.Append(CreateBarChart(plotArea, range, descriptors, direction, grouping, categoryAxisId, valueAxisId));
            } else {
                ResetAxisIds(chart, categoryAxisId, valueAxisId);
                UpdateBarChartSeries(chart, data, range, descriptors);
            }

            return chart;
        }

        private static OfficeXmlNode UpdateOrCreateLineChart(OfficeXmlNode plotArea, IReadOnlyList<SeriesDescriptor> descriptors, ExcelChartData data, ExcelChartDataRange range,
            string grouping, uint categoryAxisId, uint valueAxisId) {
            OfficeXmlNode chart = FindGroupedChart(plotArea, "lineChart", grouping, categoryAxisId, valueAxisId);

            if (chart.IsNone) {
                chart = plotArea.Append(CreateLineChart(plotArea, range, descriptors, grouping, categoryAxisId, valueAxisId));
            } else {
                EnsureGrouping(chart, grouping);
                ResetAxisIds(chart, categoryAxisId, valueAxisId);
                UpdateLineChartSeries(chart, data, range, descriptors);
            }

            return chart;
        }

        private static OfficeXmlNode UpdateOrCreateAreaChart(OfficeXmlNode plotArea, IReadOnlyList<SeriesDescriptor> descriptors, ExcelChartData data, ExcelChartDataRange range,
            string grouping, uint categoryAxisId, uint valueAxisId) {
            OfficeXmlNode chart = FindGroupedChart(plotArea, "areaChart", grouping, categoryAxisId, valueAxisId);

            if (chart.IsNone) {
                chart = plotArea.Append(CreateAreaChart(plotArea, range, descriptors, grouping, categoryAxisId, valueAxisId));
            } else {
                EnsureGrouping(chart, grouping);
                ResetAxisIds(chart, categoryAxisId, valueAxisId);
                UpdateAreaChartSeries(chart, data, range, descriptors);
            }

            return chart;
        }

        // The first line or area chart with this grouping on these axes, else the first with this grouping, else none.
        private static OfficeXmlNode FindGroupedChart(OfficeXmlNode plotArea, string localName, string grouping, uint categoryAxisId, uint valueAxisId) {
            foreach (OfficeXmlNode candidate in plotArea.Children(C, localName)) {
                if (IsGroupedChartOf(candidate, grouping) && ChartHasAxisIds(candidate, categoryAxisId, valueAxisId)) return candidate;
            }
            foreach (OfficeXmlNode candidate in plotArea.Children(C, localName)) {
                if (IsGroupedChartOf(candidate, grouping)) return candidate;
            }
            return OfficeXmlNode.None;
        }

        private static void EnsureGrouping(OfficeXmlNode chart, string groupingValue) {
            OfficeXmlNode grouping = chart.Child(C, "grouping");
            if (grouping.IsNone) grouping = chart.InsertBefore(chart.New(C, "grouping"), chart.FirstChild);
            grouping.SetVal(groupingValue);
        }

        private static OfficeXmlNode UpdateOrCreateScatterChart(OfficeXmlNode plotArea, IReadOnlyList<SeriesDescriptor> descriptors, ExcelChartData data, ExcelChartDataRange range,
            uint xAxisId, uint yAxisId) {
            OfficeXmlNode chart = OfficeXmlNode.None;
            foreach (OfficeXmlNode candidate in plotArea.Children(C, "scatterChart")) {
                if (ChartHasAxisIds(candidate, xAxisId, yAxisId)) {
                    chart = candidate;
                    break;
                }
            }
            if (chart.IsNone) chart = plotArea.Child(C, "scatterChart");

            if (chart.IsNone) {
                chart = plotArea.Append(CreateScatterChart(plotArea, range, descriptors, xAxisId, yAxisId, data));
            } else {
                ResetAxisIds(chart, xAxisId, yAxisId);
                UpdateScatterChartSeries(chart, data, range, descriptors);
            }

            return chart;
        }

        private static bool ChartHasAxisIds(OfficeXmlNode chart, uint categoryAxisId, uint valueAxisId) {
            bool hasCategoryAxis = false;
            bool hasValueAxis = false;
            foreach (OfficeXmlNode id in chart.Children(C, "axId")) {
                uint? value = id.UInt32Val();
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

        private static void ResetAxisIds(OfficeXmlNode chart, uint categoryAxisId, uint valueAxisId) {
            chart.RemoveChildren(C, "axId");
            chart.Append(NewAxisId(chart, categoryAxisId));
            chart.Append(NewAxisId(chart, valueAxisId));
        }
    }
}
