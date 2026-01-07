using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using ChartIndex = DocumentFormat.OpenXml.Drawing.Charts.Index;

namespace OfficeIMO.Excel {
    internal static partial class ExcelChartUtils {
        internal static void UpdateChartData(ChartPart chartPart, ExcelChartData data, ExcelChartDataRange range) {
            if (chartPart == null) {
                throw new ArgumentNullException(nameof(chartPart));
            }
            if (data == null) {
                throw new ArgumentNullException(nameof(data));
            }

            ChartSpace? chartSpace = chartPart.ChartSpace;
            Chart? chart = chartSpace?.GetFirstChild<Chart>();
            PlotArea? plotArea = chart?.GetFirstChild<PlotArea>();
            if (plotArea == null) {
                throw new InvalidOperationException("Chart plot area not found.");
            }

            int chartElementCount =
                plotArea.Elements<BarChart>().Count()
                + plotArea.Elements<LineChart>().Count()
                + plotArea.Elements<AreaChart>().Count()
                + plotArea.Elements<PieChart>().Count()
                + plotArea.Elements<DoughnutChart>().Count()
                + plotArea.Elements<ScatterChart>().Count()
                + plotArea.Elements<BubbleChart>().Count();

            ExcelChartType defaultType = InferChartType(plotArea);
            List<SeriesDescriptor> descriptors = BuildSeriesDescriptors(range, data, defaultType, useSeriesOverrides: chartElementCount > 1);

            if (plotArea.GetFirstChild<ScatterChart>() is ScatterChart scatterChart) {
                if (chartElementCount > 1) {
                    UpdateComboChartData(plotArea, data, range, descriptors);
                } else {
                    UpdateScatterChartSeries(scatterChart, data, range, descriptors);
                }
                return;
            }

            if (plotArea.GetFirstChild<BubbleChart>() != null) {
                throw new NotSupportedException("Updating bubble charts is not supported. Use range-based charts for bubble data.");
            }

            if (chartElementCount > 1) {
                UpdateComboChartData(plotArea, data, range, descriptors);
                return;
            }

            if (plotArea.GetFirstChild<BarChart>() is BarChart barChart) {
                UpdateBarChartSeries(barChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<LineChart>() is LineChart lineChart) {
                UpdateLineChartSeries(lineChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<AreaChart>() is AreaChart areaChart) {
                UpdateAreaChartSeries(areaChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<PieChart>() is PieChart pieChart) {
                UpdatePieChartSeries(pieChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<DoughnutChart>() is DoughnutChart doughnutChart) {
                UpdateDoughnutChartSeries(doughnutChart, data, range, descriptors);
                return;
            }

            throw new NotSupportedException("Chart type is not supported for data updates.");
        }

        internal static ExcelChartDataRange? TryExtractDataRange(ChartPart chartPart) {
            var chart = chartPart.ChartSpace?.GetFirstChild<Chart>();
            var plotArea = chart?.GetFirstChild<PlotArea>();
            if (plotArea == null) return null;

            IReadOnlyList<OpenXmlCompositeElement> seriesList;
            if (plotArea.GetFirstChild<BarChart>() is BarChart bar) {
                seriesList = bar.Elements<BarChartSeries>().Cast<OpenXmlCompositeElement>().ToList();
            } else if (plotArea.GetFirstChild<LineChart>() is LineChart line) {
                seriesList = line.Elements<LineChartSeries>().Cast<OpenXmlCompositeElement>().ToList();
            } else if (plotArea.GetFirstChild<AreaChart>() is AreaChart area) {
                seriesList = area.Elements<AreaChartSeries>().Cast<OpenXmlCompositeElement>().ToList();
            } else if (plotArea.GetFirstChild<PieChart>() is PieChart pie) {
                seriesList = pie.Elements<PieChartSeries>().Cast<OpenXmlCompositeElement>().ToList();
            } else if (plotArea.GetFirstChild<DoughnutChart>() is DoughnutChart doughnut) {
                seriesList = doughnut.Elements<PieChartSeries>().Cast<OpenXmlCompositeElement>().ToList();
            } else if (plotArea.GetFirstChild<ScatterChart>() is ScatterChart scatter) {
                seriesList = scatter.Elements<ScatterChartSeries>().Cast<OpenXmlCompositeElement>().ToList();
            } else {
                return null;
            }

            if (seriesList.Count == 0) return null;

            var series = seriesList[0];
            string? catFormula;
            string? valFormula;
            if (series is ScatterChartSeries scatterSeries) {
                catFormula = scatterSeries.GetFirstChild<XValues>()?
                    .GetFirstChild<NumberReference>()?
                    .Formula?.Text;
                valFormula = scatterSeries.GetFirstChild<YValues>()?
                    .GetFirstChild<NumberReference>()?
                    .Formula?.Text;
            } else {
                catFormula = series.GetFirstChild<CategoryAxisData>()?
                    .GetFirstChild<StringReference>()?
                    .Formula?.Text;
                valFormula = series.GetFirstChild<Values>()?
                    .GetFirstChild<NumberReference>()?
                    .Formula?.Text;
            }

            if (!TryParseSheetQualifiedRange(catFormula, out var sheetName, out var catRange)) return null;
            if (!TryParseSheetQualifiedRange(valFormula, out var sheetNameValues, out var valRange)) return null;

            if (!string.Equals(sheetName, sheetNameValues, StringComparison.OrdinalIgnoreCase)) return null;
            if (!TryParseRange(catRange, out int r1, out int c1, out int r2, out _)) return null;
            if (!TryParseRange(valRange, out _, out _, out _, out _)) return null;

            int categoryCount = r2 - r1 + 1;
            if (categoryCount <= 0) return null;

            bool hasHeaderRow = false;
            int headerRow = r1 - 1;
            if (headerRow > 0) {
                foreach (var seriesElement in seriesList) {
                    string? nameFormula = seriesElement.GetFirstChild<SeriesText>()?
                        .GetFirstChild<StringReference>()?
                        .Formula?.Text;
                    if (!TryParseSheetQualifiedRange(nameFormula, out var nameSheet, out var nameRange)) {
                        continue;
                    }
                    if (!string.Equals(nameSheet, sheetName, StringComparison.OrdinalIgnoreCase)) {
                        continue;
                    }
                    if (!TryParseRange(nameRange, out int nameR1, out _, out int nameR2, out _)) {
                        continue;
                    }
                    if (nameR1 == nameR2 && nameR1 == headerRow) {
                        hasHeaderRow = true;
                        break;
                    }
                }
            }

            int startRow = hasHeaderRow ? headerRow : r1;
            int startColumn = c1;

            return new ExcelChartDataRange(sheetName, startRow, startColumn, categoryCount, seriesList.Count, hasHeaderRow);
        }

        internal static ExcelChartData? TryReadChartData(ExcelSheet sheet, ExcelChartDataRange range) {
            try {
                var categories = new List<string>(range.CategoryCount);
                for (int i = 0; i < range.CategoryCount; i++) {
                    int row = range.CategoryStartRow + i;
                    if (sheet.TryGetCellText(row, range.StartColumn, out var text)) {
                        categories.Add(text ?? string.Empty);
                    } else {
                        categories.Add(string.Empty);
                    }
                }

                var series = new List<ExcelChartSeries>(range.SeriesCount);
                for (int s = 0; s < range.SeriesCount; s++) {
                    int col = range.SeriesStartColumn + s;
                    string name = $"Series {s + 1}";
                    if (range.HasHeaderRow) {
                        if (sheet.TryGetCellText(range.StartRow, col, out var header) && !string.IsNullOrWhiteSpace(header)) {
                            name = header;
                        }
                    }

                    var values = new List<double>(range.CategoryCount);
                    for (int i = 0; i < range.CategoryCount; i++) {
                        int row = range.CategoryStartRow + i;
                        if (sheet.TryGetCellText(row, col, out var raw)
                            && double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out var val)) {
                            values.Add(val);
                        } else {
                            values.Add(0d);
                        }
                    }
                    series.Add(new ExcelChartSeries(name, values));
                }

                return new ExcelChartData(categories, series);
            } catch {
                return null;
            }
        }

        private static void UpdateBarChartSeries(BarChart barChart, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            List<BarChartSeries> existingSeries = barChart.Elements<BarChartSeries>().ToList();
            BarChartSeries? template = existingSeries.LastOrDefault();
            var indexSet = new HashSet<int>(descriptors.Select(d => d.Index));
            var existingByIndex = existingSeries.ToDictionary(GetSeriesIndex, s => s);

            foreach (var descriptor in descriptors) {
                BarChartSeries seriesElement;
                if (!existingByIndex.TryGetValue(descriptor.Index, out seriesElement!)) {
                    seriesElement = template != null ? (BarChartSeries)template.CloneNode(true) : new BarChartSeries();
                    InsertSeries(barChart, seriesElement);
                }

                UpdateSeriesIndexOrder(seriesElement, descriptor.Index);
                string name = descriptor.Series?.Name ?? $"Series {descriptor.Index + 1}";
                UpdateSeriesText(seriesElement, range, descriptor.Index, name);
                UpdateCategoryAxisData(seriesElement, range, data.Categories);
                UpdateValues(seriesElement, range, descriptor.Index, data.Series[descriptor.Index].Values);
            }

            foreach (var seriesElement in existingSeries) {
                int idx = GetSeriesIndex(seriesElement);
                if (!indexSet.Contains(idx)) {
                    seriesElement.Remove();
                }
            }
        }

        private static void UpdateLineChartSeries(LineChart lineChart, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            List<LineChartSeries> existingSeries = lineChart.Elements<LineChartSeries>().ToList();
            LineChartSeries? template = existingSeries.LastOrDefault();
            var indexSet = new HashSet<int>(descriptors.Select(d => d.Index));
            var existingByIndex = existingSeries.ToDictionary(GetSeriesIndex, s => s);

            foreach (var descriptor in descriptors) {
                LineChartSeries seriesElement;
                if (!existingByIndex.TryGetValue(descriptor.Index, out seriesElement!)) {
                    seriesElement = template != null ? (LineChartSeries)template.CloneNode(true) : new LineChartSeries();
                    InsertSeries(lineChart, seriesElement);
                }

                UpdateSeriesIndexOrder(seriesElement, descriptor.Index);
                string name = descriptor.Series?.Name ?? $"Series {descriptor.Index + 1}";
                UpdateSeriesText(seriesElement, range, descriptor.Index, name);
                UpdateCategoryAxisData(seriesElement, range, data.Categories);
                UpdateValues(seriesElement, range, descriptor.Index, data.Series[descriptor.Index].Values);
            }

            foreach (var seriesElement in existingSeries) {
                int idx = GetSeriesIndex(seriesElement);
                if (!indexSet.Contains(idx)) {
                    seriesElement.Remove();
                }
            }
        }

        private static void UpdateAreaChartSeries(AreaChart areaChart, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            List<AreaChartSeries> existingSeries = areaChart.Elements<AreaChartSeries>().ToList();
            AreaChartSeries? template = existingSeries.LastOrDefault();
            var indexSet = new HashSet<int>(descriptors.Select(d => d.Index));
            var existingByIndex = existingSeries.ToDictionary(GetSeriesIndex, s => s);

            foreach (var descriptor in descriptors) {
                AreaChartSeries seriesElement;
                if (!existingByIndex.TryGetValue(descriptor.Index, out seriesElement!)) {
                    seriesElement = template != null ? (AreaChartSeries)template.CloneNode(true) : new AreaChartSeries();
                    InsertSeries(areaChart, seriesElement);
                }

                UpdateSeriesIndexOrder(seriesElement, descriptor.Index);
                string name = descriptor.Series?.Name ?? $"Series {descriptor.Index + 1}";
                UpdateSeriesText(seriesElement, range, descriptor.Index, name);
                UpdateCategoryAxisData(seriesElement, range, data.Categories);
                UpdateValues(seriesElement, range, descriptor.Index, data.Series[descriptor.Index].Values);
            }

            foreach (var seriesElement in existingSeries) {
                int idx = GetSeriesIndex(seriesElement);
                if (!indexSet.Contains(idx)) {
                    seriesElement.Remove();
                }
            }
        }

        private static void UpdatePieChartSeries(PieChart chart, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            List<PieChartSeries> existingSeries = chart.Elements<PieChartSeries>().ToList();
            PieChartSeries? template = existingSeries.LastOrDefault();
            var indexSet = new HashSet<int>(descriptors.Select(d => d.Index));
            var existingByIndex = existingSeries.ToDictionary(GetSeriesIndex, s => s);

            foreach (var descriptor in descriptors) {
                PieChartSeries seriesElement;
                if (!existingByIndex.TryGetValue(descriptor.Index, out seriesElement!)) {
                    seriesElement = template != null ? (PieChartSeries)template.CloneNode(true) : new PieChartSeries();
                    InsertSeries(chart, seriesElement);
                }

                UpdateSeriesIndexOrder(seriesElement, descriptor.Index);
                string name = descriptor.Series?.Name ?? $"Series {descriptor.Index + 1}";
                UpdateSeriesText(seriesElement, range, descriptor.Index, name);
                UpdateCategoryAxisData(seriesElement, range, data.Categories);
                UpdateValues(seriesElement, range, descriptor.Index, data.Series[descriptor.Index].Values);
            }

            foreach (var seriesElement in existingSeries) {
                int idx = GetSeriesIndex(seriesElement);
                if (!indexSet.Contains(idx)) {
                    seriesElement.Remove();
                }
            }
        }

        private static void UpdateDoughnutChartSeries(DoughnutChart chart, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            List<PieChartSeries> existingSeries = chart.Elements<PieChartSeries>().ToList();
            PieChartSeries? template = existingSeries.LastOrDefault();
            var indexSet = new HashSet<int>(descriptors.Select(d => d.Index));
            var existingByIndex = existingSeries.ToDictionary(GetSeriesIndex, s => s);

            foreach (var descriptor in descriptors) {
                PieChartSeries seriesElement;
                if (!existingByIndex.TryGetValue(descriptor.Index, out seriesElement!)) {
                    seriesElement = template != null ? (PieChartSeries)template.CloneNode(true) : new PieChartSeries();
                    InsertSeries(chart, seriesElement);
                }

                UpdateSeriesIndexOrder(seriesElement, descriptor.Index);
                string name = descriptor.Series?.Name ?? $"Series {descriptor.Index + 1}";
                UpdateSeriesText(seriesElement, range, descriptor.Index, name);
                UpdateCategoryAxisData(seriesElement, range, data.Categories);
                UpdateValues(seriesElement, range, descriptor.Index, data.Series[descriptor.Index].Values);
            }

            foreach (var seriesElement in existingSeries) {
                int idx = GetSeriesIndex(seriesElement);
                if (!indexSet.Contains(idx)) {
                    seriesElement.Remove();
                }
            }
        }

        private static void UpdateScatterChartSeries(ScatterChart chart, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            List<ScatterChartSeries> existingSeries = chart.Elements<ScatterChartSeries>().ToList();
            ScatterChartSeries? template = existingSeries.LastOrDefault();
            var indexSet = new HashSet<int>(descriptors.Select(d => d.Index));
            var existingByIndex = existingSeries.ToDictionary(GetSeriesIndex, s => s);
            IReadOnlyList<double> xValues = ParseNumericCategories(data.Categories);

            foreach (var descriptor in descriptors) {
                ScatterChartSeries seriesElement;
                if (!existingByIndex.TryGetValue(descriptor.Index, out seriesElement!)) {
                    seriesElement = template != null ? (ScatterChartSeries)template.CloneNode(true) : new ScatterChartSeries();
                    InsertSeries(chart, seriesElement);
                }

                UpdateSeriesIndexOrder(seriesElement, descriptor.Index);
                string name = descriptor.Series?.Name ?? $"Series {descriptor.Index + 1}";
                UpdateSeriesText(seriesElement, range, descriptor.Index, name);
                UpdateXValues(seriesElement, range, xValues);
                UpdateYValues(seriesElement, range, descriptor.Index, data.Series[descriptor.Index].Values);
            }

            foreach (var seriesElement in existingSeries) {
                int idx = GetSeriesIndex(seriesElement);
                if (!indexSet.Contains(idx)) {
                    seriesElement.Remove();
                }
            }
        }

        private static void UpdateComboChartData(PlotArea plotArea, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            if (plotArea == null) {
                throw new ArgumentNullException(nameof(plotArea));
            }

            bool hasSecondary = descriptors.Any(d => d.AxisGroup == ExcelChartAxisGroup.Secondary);
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

            var axisIds = EnsureAxisPairs(plotArea, hasSecondary, primaryCategoryPosition, primaryValuePosition, secondaryCategoryPosition, secondaryValuePosition);

            var usedCharts = new List<OpenXmlCompositeElement>();
            foreach (var group in descriptors.GroupBy(d => new { d.ChartType, d.AxisGroup })) {
                uint categoryAxisId = group.Key.AxisGroup == ExcelChartAxisGroup.Secondary ? axisIds.SecondaryCategoryId : axisIds.PrimaryCategoryId;
                uint valueAxisId = group.Key.AxisGroup == ExcelChartAxisGroup.Secondary ? axisIds.SecondaryValueId : axisIds.PrimaryValueId;
                var groupDescriptors = group.ToList();

                switch (group.Key.ChartType) {
                    case ExcelChartType.ColumnClustered:
                        usedCharts.Add(UpdateOrCreateBarChart(plotArea, groupDescriptors, data, range, BarDirectionValues.Column, BarGroupingValues.Clustered, categoryAxisId, valueAxisId));
                        break;
                    case ExcelChartType.ColumnStacked:
                        usedCharts.Add(UpdateOrCreateBarChart(plotArea, groupDescriptors, data, range, BarDirectionValues.Column, BarGroupingValues.Stacked, categoryAxisId, valueAxisId));
                        break;
                    case ExcelChartType.BarClustered:
                        usedCharts.Add(UpdateOrCreateBarChart(plotArea, groupDescriptors, data, range, BarDirectionValues.Bar, BarGroupingValues.Clustered, categoryAxisId, valueAxisId));
                        break;
                    case ExcelChartType.BarStacked:
                        usedCharts.Add(UpdateOrCreateBarChart(plotArea, groupDescriptors, data, range, BarDirectionValues.Bar, BarGroupingValues.Stacked, categoryAxisId, valueAxisId));
                        break;
                    case ExcelChartType.Line:
                        usedCharts.Add(UpdateOrCreateLineChart(plotArea, groupDescriptors, data, range, categoryAxisId, valueAxisId));
                        break;
                    case ExcelChartType.Area:
                        usedCharts.Add(UpdateOrCreateAreaChart(plotArea, groupDescriptors, data, range, categoryAxisId, valueAxisId));
                        break;
                    case ExcelChartType.Scatter:
                        usedCharts.Add(UpdateOrCreateScatterChart(plotArea, groupDescriptors, data, range, categoryAxisId, valueAxisId));
                        break;
                    default:
                        throw new NotSupportedException($"Chart type {group.Key.ChartType} is not supported in combination charts.");
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
            uint categoryAxisId, uint valueAxisId) {
            LineChart? chart = plotArea.Elements<LineChart>()
                .FirstOrDefault(c => ChartHasAxisIds(c, categoryAxisId, valueAxisId))
                ?? plotArea.Elements<LineChart>().FirstOrDefault();

            if (chart == null) {
                chart = CreateLineChart(range, descriptors, categoryAxisId, valueAxisId);
                plotArea.Append(chart);
            } else {
                ResetAxisIds(chart, categoryAxisId, valueAxisId);
                UpdateLineChartSeries(chart, data, range, descriptors);
            }

            return chart;
        }

        private static AreaChart UpdateOrCreateAreaChart(PlotArea plotArea, IReadOnlyList<SeriesDescriptor> descriptors, ExcelChartData data, ExcelChartDataRange range,
            uint categoryAxisId, uint valueAxisId) {
            AreaChart? chart = plotArea.Elements<AreaChart>()
                .FirstOrDefault(c => ChartHasAxisIds(c, categoryAxisId, valueAxisId))
                ?? plotArea.Elements<AreaChart>().FirstOrDefault();

            if (chart == null) {
                chart = CreateAreaChart(range, descriptors, categoryAxisId, valueAxisId);
                plotArea.Append(chart);
            } else {
                ResetAxisIds(chart, categoryAxisId, valueAxisId);
                UpdateAreaChartSeries(chart, data, range, descriptors);
            }

            return chart;
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
            var ids = chart.Elements<AxisId>()
                .Select(id => id.Val?.Value)
                .Where(val => val.HasValue)
                .Select(val => val!.Value)
                .ToList();
            return ids.Contains(categoryAxisId) && ids.Contains(valueAxisId);
        }

        private static void ResetAxisIds(OpenXmlCompositeElement chart, uint categoryAxisId, uint valueAxisId) {
            chart.RemoveAllChildren<AxisId>();
            chart.Append(new AxisId { Val = categoryAxisId });
            chart.Append(new AxisId { Val = valueAxisId });
        }

        private static void UpdateSeriesIndexOrder(OpenXmlCompositeElement series, int index) {
            ChartIndex idx = series.GetFirstChild<ChartIndex>() ?? new ChartIndex();
            idx.Val = (uint)index;
            if (idx.Parent == null) {
                series.PrependChild(idx);
            }

            Order order = series.GetFirstChild<Order>() ?? new Order();
            order.Val = (uint)index;
            if (order.Parent == null) {
                series.InsertAfter(order, idx);
            }
        }

        private static int GetSeriesIndex(OpenXmlCompositeElement series) {
            return (int)(series.GetFirstChild<ChartIndex>()?.Val?.Value ?? 0U);
        }

        private static void UpdateSeriesText(OpenXmlCompositeElement series, ExcelChartDataRange range, int seriesIndex, string seriesName) {
            SeriesText seriesText = series.GetFirstChild<SeriesText>() ?? new SeriesText();
            seriesText.RemoveAllChildren<StringReference>();
            seriesText.RemoveAllChildren<StringLiteral>();

            if (range.HasHeaderRow) {
                string seriesCell = range.SeriesNameCellA1(seriesIndex);
                string formula = BuildSheetQualifiedRange(range.SheetName, seriesCell);
                seriesText.Append(CreateStringReference(formula, new[] { seriesName }));
            } else {
                seriesText.Append(CreateStringLiteral(new[] { seriesName }));
            }

            if (seriesText.Parent == null) {
                OpenXmlElement? insertAfter = series.GetFirstChild<Order>();
                insertAfter ??= series.GetFirstChild<ChartIndex>();
                if (insertAfter != null) {
                    series.InsertAfter(seriesText, insertAfter);
                } else {
                    series.PrependChild(seriesText);
                }
            }
        }

        private static void UpdateCategoryAxisData(OpenXmlCompositeElement series, ExcelChartDataRange range, IReadOnlyList<string> categories) {
            string formula = BuildSheetQualifiedRange(range.SheetName, range.CategoriesRangeA1);
            CategoryAxisData categoryAxisData = series.GetFirstChild<CategoryAxisData>() ?? new CategoryAxisData();
            categoryAxisData.RemoveAllChildren<StringReference>();
            categoryAxisData.RemoveAllChildren<StringLiteral>();
            categoryAxisData.Append(CreateStringReference(formula, categories));

            if (categoryAxisData.Parent == null) {
                series.Append(categoryAxisData);
            }
        }

        private static void UpdateValues(OpenXmlCompositeElement series, ExcelChartDataRange range, int seriesIndex, IReadOnlyList<double> values) {
            string formula = BuildSheetQualifiedRange(range.SheetName, range.SeriesValuesRangeA1(seriesIndex));
            Values valueElement = series.GetFirstChild<Values>() ?? new Values();
            valueElement.RemoveAllChildren<NumberReference>();
            valueElement.RemoveAllChildren<NumberLiteral>();
            valueElement.Append(CreateNumberReference(formula, values));

            if (valueElement.Parent == null) {
                series.Append(valueElement);
            }
        }

        private static void UpdateXValues(ScatterChartSeries series, ExcelChartDataRange range, IReadOnlyList<double> xValues) {
            string formula = BuildSheetQualifiedRange(range.SheetName, range.CategoriesRangeA1);
            XValues xValueElement = series.GetFirstChild<XValues>() ?? new XValues();
            xValueElement.RemoveAllChildren<NumberReference>();
            xValueElement.RemoveAllChildren<NumberLiteral>();
            xValueElement.Append(CreateNumberReference(formula, xValues));

            if (xValueElement.Parent == null) {
                series.Append(xValueElement);
            }
        }

        private static void UpdateYValues(ScatterChartSeries series, ExcelChartDataRange range, int seriesIndex, IReadOnlyList<double> values) {
            string formula = BuildSheetQualifiedRange(range.SheetName, range.SeriesValuesRangeA1(seriesIndex));
            YValues yValueElement = series.GetFirstChild<YValues>() ?? new YValues();
            yValueElement.RemoveAllChildren<NumberReference>();
            yValueElement.RemoveAllChildren<NumberLiteral>();
            yValueElement.Append(CreateNumberReference(formula, values));

            if (yValueElement.Parent == null) {
                series.Append(yValueElement);
            }
        }

        private static void InsertSeries(OpenXmlCompositeElement chart, OpenXmlElement series) {
            OpenXmlElement? insertBefore = chart.ChildElements.FirstOrDefault(child =>
                child is DataLabels ||
                child is GapWidth ||
                child is Overlap ||
                child is AxisId ||
                child is Marker ||
                child is Smooth ||
                child is SeriesLines);

            if (insertBefore != null) {
                chart.InsertBefore(series, insertBefore);
            } else {
                chart.Append(series);
            }
        }
    }
}
