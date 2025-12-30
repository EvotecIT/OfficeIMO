using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using ChartIndex = DocumentFormat.OpenXml.Drawing.Charts.Index;

namespace OfficeIMO.Excel {
    internal static class ExcelChartUtils {
        private const string ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        private const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly Lazy<byte[]> ChartStyle251Bytes =
            new(() => LoadEmbeddedResource("OfficeIMO.Excel.Resources.chart-style-251.xml"));
        private static readonly Lazy<byte[]> ChartColorStyle10Bytes =
            new(() => LoadEmbeddedResource("OfficeIMO.Excel.Resources.chart-colors-10.xml"));

        private sealed class SeriesDescriptor {
            public SeriesDescriptor(int index, ExcelChartSeries? series, ExcelChartType chartType, ExcelChartAxisGroup axisGroup) {
                Index = index;
                Series = series;
                ChartType = chartType;
                AxisGroup = axisGroup;
            }

            public int Index { get; }
            public ExcelChartSeries? Series { get; }
            public ExcelChartType ChartType { get; }
            public ExcelChartAxisGroup AxisGroup { get; }
        }

        private readonly struct AxisIdSet {
            public AxisIdSet(uint primaryCategoryId, uint primaryValueId, uint secondaryCategoryId, uint secondaryValueId) {
                PrimaryCategoryId = primaryCategoryId;
                PrimaryValueId = primaryValueId;
                SecondaryCategoryId = secondaryCategoryId;
                SecondaryValueId = secondaryValueId;
            }

            public uint PrimaryCategoryId { get; }
            public uint PrimaryValueId { get; }
            public uint SecondaryCategoryId { get; }
            public uint SecondaryValueId { get; }
        }

        internal static string BuildCellA1(int row, int column) {
            string col = A1.ColumnIndexToLetters(column);
            return $"${col}${row}";
        }

        internal static string BuildRangeA1(int row1, int col1, int row2, int col2) {
            string start = BuildCellA1(row1, col1);
            string end = BuildCellA1(row2, col2);
            return $"{start}:{end}";
        }

        internal static string BuildSheetQualifiedRange(string sheetName, string a1Range) {
            return $"{QuoteSheetName(sheetName)}!{a1Range}";
        }

        internal static string EnsureSheetQualifiedRange(string sheetName, string a1Range) {
            if (string.IsNullOrWhiteSpace(a1Range)) {
                throw new ArgumentException("Range cannot be null or empty.", nameof(a1Range));
            }

            string trimmed = a1Range.Trim();
            return trimmed.Contains("!") ? trimmed : BuildSheetQualifiedRange(sheetName, trimmed);
        }

        internal static string QuoteSheetName(string sheetName) {
            var escaped = sheetName.Replace("'", "''");
            return $"'{escaped}'";
        }

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

            bool hasHeaderRow = r1 > 1;
            int startRow = hasHeaderRow ? r1 - 1 : r1;
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

        internal static bool TryParseSheetQualifiedRange(string? formula, out string sheetName, out string rangeA1) {
            sheetName = string.Empty;
            rangeA1 = string.Empty;
            if (string.IsNullOrWhiteSpace(formula)) return false;

            string formulaValue = formula ?? string.Empty;
            int bang = formulaValue.LastIndexOf('!');
            if (bang <= 0 || bang >= formulaValue.Length - 1) return false;

            sheetName = formulaValue.Substring(0, bang);
            rangeA1 = formulaValue.Substring(bang + 1);

            sheetName = UnquoteSheetName(sheetName);
            rangeA1 = rangeA1.Replace("$", string.Empty);
            return true;
        }

        private static bool TryParseRange(string a1Range, out int r1, out int c1, out int r2, out int c2) {
            r1 = c1 = r2 = c2 = 0;
            if (string.IsNullOrWhiteSpace(a1Range)) return false;
            string normalized = a1Range.Replace("$", string.Empty);
            if (normalized.Contains(":")) {
                return A1.TryParseRange(normalized, out r1, out c1, out r2, out c2);
            }
            var (row, col) = A1.ParseCellRef(normalized);
            if (row <= 0 || col <= 0) return false;
            r1 = r2 = row;
            c1 = c2 = col;
            return true;
        }

        private static string UnquoteSheetName(string name) {
            name = name.Trim();
            if (name.Length >= 2 && name[0] == '\'' && name[name.Length - 1] == '\'') {
                name = name.Substring(1, name.Length - 2).Replace("''", "'");
            }
            return name;
        }

        private static List<SeriesDescriptor> BuildSeriesDescriptors(ExcelChartDataRange range, ExcelChartData? data, ExcelChartType defaultType, bool useSeriesOverrides = true) {
            int count = data?.Series.Count ?? range.SeriesCount;
            var descriptors = new List<SeriesDescriptor>(count);
            for (int i = 0; i < count; i++) {
                var series = data?.Series.ElementAtOrDefault(i);
                ExcelChartType chartType = defaultType;
                ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary;
                if (useSeriesOverrides && series != null) {
                    chartType = series.ChartType ?? defaultType;
                    axisGroup = series.AxisGroup;
                }
                descriptors.Add(new SeriesDescriptor(i, series, chartType, axisGroup));
            }
            return descriptors;
        }

        private static ExcelChartType InferChartType(PlotArea plotArea) {
            if (plotArea.GetFirstChild<BarChart>() is BarChart barChart) {
                BarDirectionValues direction = barChart.GetFirstChild<BarDirection>()?.Val ?? BarDirectionValues.Column;
                BarGroupingValues grouping = barChart.GetFirstChild<BarGrouping>()?.Val ?? BarGroupingValues.Clustered;
                if (direction == BarDirectionValues.Bar) {
                    return grouping == BarGroupingValues.Stacked ? ExcelChartType.BarStacked : ExcelChartType.BarClustered;
                }
                return grouping == BarGroupingValues.Stacked ? ExcelChartType.ColumnStacked : ExcelChartType.ColumnClustered;
            }
            if (plotArea.GetFirstChild<LineChart>() != null) return ExcelChartType.Line;
            if (plotArea.GetFirstChild<AreaChart>() != null) return ExcelChartType.Area;
            if (plotArea.GetFirstChild<PieChart>() != null) return ExcelChartType.Pie;
            if (plotArea.GetFirstChild<DoughnutChart>() != null) return ExcelChartType.Doughnut;
            if (plotArea.GetFirstChild<ScatterChart>() != null) return ExcelChartType.Scatter;
            if (plotArea.GetFirstChild<BubbleChart>() != null) return ExcelChartType.Bubble;
            return ExcelChartType.ColumnClustered;
        }

        private static bool IsBarChartType(ExcelChartType chartType) {
            return chartType == ExcelChartType.BarClustered || chartType == ExcelChartType.BarStacked;
        }

        private static (BarDirectionValues Direction, BarGroupingValues Grouping) GetBarChartSettings(ExcelChartType chartType) {
            return chartType switch {
                ExcelChartType.ColumnClustered => (BarDirectionValues.Column, BarGroupingValues.Clustered),
                ExcelChartType.ColumnStacked => (BarDirectionValues.Column, BarGroupingValues.Stacked),
                ExcelChartType.BarClustered => (BarDirectionValues.Bar, BarGroupingValues.Clustered),
                ExcelChartType.BarStacked => (BarDirectionValues.Bar, BarGroupingValues.Stacked),
                _ => (BarDirectionValues.Column, BarGroupingValues.Clustered)
            };
        }

        private static IReadOnlyList<double> ParseNumericCategories(IReadOnlyList<string> categories) {
            var values = new double[categories.Count];
            for (int i = 0; i < categories.Count; i++) {
                if (!double.TryParse(categories[i], NumberStyles.Any, CultureInfo.InvariantCulture, out var val)) {
                    val = 0d;
                }
                values[i] = val;
            }
            return values;
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

        private static BarChart CreateBarChart(ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors, BarDirectionValues direction,
            BarGroupingValues grouping, uint categoryAxisId, uint valueAxisId) {
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

        private static LineChart CreateLineChart(ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors, uint categoryAxisId, uint valueAxisId) {
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

        private static AreaChart CreateAreaChart(ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors, uint categoryAxisId, uint valueAxisId) {
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

        private static ScatterChart CreateScatterChart(ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors, uint xAxisId, uint yAxisId, ExcelChartData? data) {
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

        private static ScatterChartSeries CreateScatterChartSeries(int seriesIndex, ExcelChartDataRange range, ExcelChartSeries? series, IReadOnlyList<double>? xValues) {
            return new ScatterChartSeries(
                new ChartIndex { Val = (uint)seriesIndex },
                new Order { Val = (uint)seriesIndex },
                CreateSeriesText(range, seriesIndex, series?.Name ?? $"Series {seriesIndex + 1}"),
                CreateXValues(range, xValues),
                CreateYValues(range, seriesIndex, series)
            );
        }

        private static ScatterChart CreateScatterChartFromRanges(IReadOnlyList<ExcelChartSeriesRange> seriesRanges, string defaultSheetName, uint xAxisId, uint yAxisId) {
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

        private static ScatterChartSeries CreateScatterChartSeriesFromRanges(int seriesIndex, ExcelChartSeriesRange seriesRange, string defaultSheetName) {
            string name = string.IsNullOrWhiteSpace(seriesRange.Name) ? $"Series {seriesIndex + 1}" : seriesRange.Name;
            return new ScatterChartSeries(
                new ChartIndex { Val = (uint)seriesIndex },
                new Order { Val = (uint)seriesIndex },
                CreateSeriesTextLiteral(name),
                CreateXValuesFromRange(seriesRange.XRangeA1, defaultSheetName),
                CreateYValuesFromRange(seriesRange.YRangeA1, defaultSheetName)
            );
        }

        private static BubbleChart CreateBubbleChartFromRanges(IReadOnlyList<ExcelChartSeriesRange> seriesRanges, string defaultSheetName, uint xAxisId, uint yAxisId) {
            var bubbleChart = new BubbleChart(
                new VaryColors { Val = false },
                new BubbleScale { Val = 100 },
                new ShowNegativeBubbles { Val = false },
                new SizeRepresents { Val = SizeRepresentsValues.Area });

            for (int i = 0; i < seriesRanges.Count; i++) {
                bubbleChart.Append(CreateBubbleChartSeriesFromRanges(i, seriesRanges[i], defaultSheetName));
            }

            bubbleChart.Append(CreateDefaultDataLabels());
            bubbleChart.Append(new AxisId { Val = xAxisId });
            bubbleChart.Append(new AxisId { Val = yAxisId });
            return bubbleChart;
        }

        private static BubbleChartSeries CreateBubbleChartSeriesFromRanges(int seriesIndex, ExcelChartSeriesRange seriesRange, string defaultSheetName) {
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

            return new SeriesText(CreateStringLiteral(new[] { name }));
        }

        private static SeriesText CreateSeriesTextLiteral(string name) {
            return new SeriesText(CreateStringLiteral(new[] { name }));
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

        private static StringReference CreateStringReference(string formula, IReadOnlyList<string>? values) {
            var reference = new StringReference(new Formula { Text = formula });
            if (values == null) return reference;

            StringCache cache = new();
            cache.Append(new PointCount { Val = (uint)values.Count });
            for (int i = 0; i < values.Count; i++) {
                cache.Append(new StringPoint {
                    Index = (uint)i,
                    NumericValue = new NumericValue { Text = values[i] ?? string.Empty }
                });
            }
            reference.Append(cache);
            return reference;
        }

        private static StringLiteral CreateStringLiteral(IReadOnlyList<string> values) {
            StringLiteral literal = new();
            literal.Append(new PointCount { Val = (uint)values.Count });
            for (int i = 0; i < values.Count; i++) {
                literal.Append(new StringPoint {
                    Index = (uint)i,
                    NumericValue = new NumericValue { Text = values[i] ?? string.Empty }
                });
            }
            return literal;
        }

        private static NumberReference CreateNumberReference(string formula, IReadOnlyList<double>? values) {
            var reference = new NumberReference(new Formula { Text = formula });
            if (values == null) return reference;

            NumberingCache cache = new();
            cache.Append(new FormatCode { Text = "General" });
            cache.Append(new PointCount { Val = (uint)values.Count });
            for (int i = 0; i < values.Count; i++) {
                cache.Append(new NumericPoint {
                    Index = (uint)i,
                    NumericValue = new NumericValue { Text = values[i].ToString(CultureInfo.InvariantCulture) }
                });
            }
            reference.Append(cache);
            return reference;
        }

        private static Title CreateChartTitle(string text) {
            return new Title(
                new ChartText(CreateChartText(text)),
                new Overlay { Val = false }
            );
        }

        private static RichText CreateChartText(string text) {
            return new RichText(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(
                    new A.ParagraphProperties(new A.DefaultRunProperties()),
                    new A.Run(new A.Text { Text = text })
                ));
        }

        internal static void ApplyChartStyle(ChartPart chartPart, int styleId, int colorStyleId) {
            if (chartPart == null) {
                throw new ArgumentNullException(nameof(chartPart));
            }
            ApplyChartStyle(chartPart, new ExcelChartStylePreset(styleId, colorStyleId));
        }

        internal static void ApplyChartStyle(ChartPart chartPart, ExcelChartStylePreset preset) {
            if (chartPart == null) {
                throw new ArgumentNullException(nameof(chartPart));
            }
            if (preset == null) {
                throw new ArgumentNullException(nameof(preset));
            }

            byte[] styleBytes = preset.StyleXmlBytes ?? GetChartStyleBytes(preset.StyleId);
            byte[] colorBytes = preset.ColorXmlBytes ?? GetChartColorStyleBytes(preset.ColorStyleId);

            ChartStylePart stylePart = chartPart.GetPartsOfType<ChartStylePart>().FirstOrDefault()
                ?? chartPart.AddNewPart<ChartStylePart>();
            PopulateChartStyle(stylePart, styleBytes);
            ChartColorStylePart colorStylePart = chartPart.GetPartsOfType<ChartColorStylePart>().FirstOrDefault()
                ?? chartPart.AddNewPart<ChartColorStylePart>();
            PopulateChartColorStyle(colorStylePart, colorBytes);
        }

        internal static void PopulateChartStyle(ChartStylePart stylePart, byte[]? xmlBytes = null) {
            if (stylePart == null) {
                throw new ArgumentNullException(nameof(stylePart));
            }

            xmlBytes ??= ChartStyle251Bytes.Value;
            using var stream = new MemoryStream(xmlBytes);
            stylePart.FeedData(stream);
        }

        internal static void PopulateChartColorStyle(ChartColorStylePart colorStylePart, byte[]? xmlBytes = null) {
            if (colorStylePart == null) {
                throw new ArgumentNullException(nameof(colorStylePart));
            }

            xmlBytes ??= ChartColorStyle10Bytes.Value;
            using var stream = new MemoryStream(xmlBytes);
            colorStylePart.FeedData(stream);
        }

        private static byte[] GetChartStyleBytes(int styleId) {
            if (styleId == 251) return ChartStyle251Bytes.Value;
            return ChartStyle251Bytes.Value;
        }

        private static byte[] GetChartColorStyleBytes(int colorStyleId) {
            if (colorStyleId == 10) return ChartColorStyle10Bytes.Value;
            return ChartColorStyle10Bytes.Value;
        }

        private static byte[] LoadEmbeddedResource(string resourceName) {
            var assembly = typeof(ExcelChartUtils).Assembly;
            using Stream? stream = assembly.GetManifestResourceStream(resourceName);
            if (stream == null) {
                throw new InvalidOperationException($"Missing embedded resource '{resourceName}'.");
            }

            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            return buffer.ToArray();
        }
    }
}
