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
        private const string ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        private const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

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

        private static List<SeriesDescriptor> BuildSeriesDescriptors(ExcelChartDataRange range, ExcelChartData? data,
            ExcelChartType defaultType, bool useSeriesOverrides = true) {
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
    }
}
