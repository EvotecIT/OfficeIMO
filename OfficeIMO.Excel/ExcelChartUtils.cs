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
            public SeriesDescriptor(int index, ExcelChartSeries? series, OfficeChartType chartType, OfficeChartAxisGroup axisGroup) {
                Index = index;
                Series = series;
                ChartType = chartType;
                AxisGroup = axisGroup;
            }

            public int Index { get; }
            public ExcelChartSeries? Series { get; }
            public OfficeChartType ChartType { get; }
            public OfficeChartAxisGroup AxisGroup { get; }
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

        private readonly struct SeriesDescriptorSummary {
            public SeriesDescriptorSummary(
                bool hasSecondary,
                bool hasScatter,
                bool hasBubble,
                bool hasStock,
                bool hasSurface,
                bool hasLine3D,
                bool hasBar3D,
                bool hasArea3D,
                bool hasBar,
                bool hasNonBar,
                bool hasPieOrDoughnut,
                bool hasMultipleTypes) {
                HasSecondary = hasSecondary;
                HasScatter = hasScatter;
                HasBubble = hasBubble;
                HasStock = hasStock;
                HasSurface = hasSurface;
                HasLine3D = hasLine3D;
                HasBar3D = hasBar3D;
                HasArea3D = hasArea3D;
                HasBar = hasBar;
                HasNonBar = hasNonBar;
                HasPieOrDoughnut = hasPieOrDoughnut;
                HasMultipleTypes = hasMultipleTypes;
            }

            public bool HasSecondary { get; }
            public bool HasScatter { get; }
            public bool HasBubble { get; }
            public bool HasStock { get; }
            public bool HasSurface { get; }
            public bool HasLine3D { get; }
            public bool HasBar3D { get; }
            public bool HasArea3D { get; }
            public bool HasBar { get; }
            public bool HasNonBar { get; }
            public bool HasPieOrDoughnut { get; }
            public bool HasMultipleTypes { get; }
        }

        private sealed class SeriesDescriptorGroup {
            public SeriesDescriptorGroup(OfficeChartType chartType, OfficeChartAxisGroup axisGroup) {
                ChartType = chartType;
                AxisGroup = axisGroup;
                Descriptors = new List<SeriesDescriptor>();
            }

            public OfficeChartType ChartType { get; }
            public OfficeChartAxisGroup AxisGroup { get; }
            public List<SeriesDescriptor> Descriptors { get; }
        }

        internal static string BuildCellA1(int row, int column) {
            return ExcelA1.AbsoluteCellReference(row, column);
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
            if (!SheetNameLookup.TryParseSheetQualifiedReference(formula, out sheetName, out rangeA1)) return false;
            rangeA1 = rangeA1.Replace("$", string.Empty);
            return true;
        }

        private static bool TryParseRange(string a1Range, out int r1, out int c1, out int r2, out int c2) {
            r1 = c1 = r2 = c2 = 0;
            if (string.IsNullOrWhiteSpace(a1Range)) return false;
            string normalized = a1Range.Replace("$", string.Empty);
            if (normalized.Contains(":")) {
                return ExcelA1.TryParseRange(normalized, out r1, out c1, out r2, out c2);
            }
            var (row, col) = ExcelA1.ParseCellRef(normalized);
            if (row <= 0 || col <= 0) return false;
            r1 = r2 = row;
            c1 = c2 = col;
            return true;
        }

        private static List<SeriesDescriptor> BuildSeriesDescriptors(ExcelChartDataRange range, ExcelChartData? data,
            OfficeChartType defaultType, bool useSeriesOverrides = true) {
            IReadOnlyList<ExcelChartSeries>? seriesList = data?.Series;
            int count = seriesList?.Count ?? range.SeriesCount;
            var descriptors = new List<SeriesDescriptor>(count);
            for (int i = 0; i < count; i++) {
                ExcelChartSeries? series = seriesList != null && i < seriesList.Count ? seriesList[i] : null;
                OfficeChartType chartType = defaultType;
                OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary;
                if (useSeriesOverrides && series != null) {
                    chartType = series.ChartType ?? defaultType;
                    axisGroup = series.AxisGroup;
                }
                descriptors.Add(new SeriesDescriptor(i, series, chartType, axisGroup));
            }
            return descriptors;
        }

        internal static OfficeChartType InferChartType(PlotArea plotArea) {
            if (plotArea.GetFirstChild<BarChart>() is BarChart barChart) {
                BarDirectionValues direction = barChart.GetFirstChild<BarDirection>()?.Val ?? BarDirectionValues.Column;
                BarGroupingValues grouping = barChart.GetFirstChild<BarGrouping>()?.Val ?? BarGroupingValues.Clustered;
                if (direction == BarDirectionValues.Bar) {
                    if (grouping == BarGroupingValues.PercentStacked) return OfficeChartType.BarStacked100;
                    return grouping == BarGroupingValues.Stacked ? OfficeChartType.BarStacked : OfficeChartType.BarClustered;
                }
                if (grouping == BarGroupingValues.PercentStacked) return OfficeChartType.ColumnStacked100;
                return grouping == BarGroupingValues.Stacked ? OfficeChartType.ColumnStacked : OfficeChartType.ColumnClustered;
            }
            if (plotArea.GetFirstChild<Bar3DChart>() is Bar3DChart bar3DChart) {
                BarDirectionValues direction = bar3DChart.GetFirstChild<BarDirection>()?.Val ?? BarDirectionValues.Column;
                BarGroupingValues grouping = bar3DChart.GetFirstChild<BarGrouping>()?.Val ?? BarGroupingValues.Clustered;
                if (direction == BarDirectionValues.Bar) {
                    if (grouping == BarGroupingValues.PercentStacked) return OfficeChartType.Bar3DStacked100;
                    return grouping == BarGroupingValues.Stacked ? OfficeChartType.Bar3DStacked : OfficeChartType.Bar3DClustered;
                }
                if (grouping == BarGroupingValues.PercentStacked) return OfficeChartType.Column3DStacked100;
                return grouping == BarGroupingValues.Stacked ? OfficeChartType.Column3DStacked : OfficeChartType.Column3DClustered;
            }
            if (plotArea.GetFirstChild<Line3DChart>() != null) return OfficeChartType.Line3D;
            if (plotArea.GetFirstChild<LineChart>() is LineChart lineChart) {
                GroupingValues grouping = lineChart.GetFirstChild<Grouping>()?.Val ?? GroupingValues.Standard;
                if (grouping == GroupingValues.PercentStacked) return OfficeChartType.LineStacked100;
                return grouping == GroupingValues.Stacked ? OfficeChartType.LineStacked : OfficeChartType.Line;
            }
            if (plotArea.GetFirstChild<Area3DChart>() is Area3DChart area3DChart) {
                GroupingValues grouping = area3DChart.GetFirstChild<Grouping>()?.Val ?? GroupingValues.Standard;
                if (grouping == GroupingValues.PercentStacked) return OfficeChartType.Area3DStacked100;
                return grouping == GroupingValues.Stacked ? OfficeChartType.Area3DStacked : OfficeChartType.Area3D;
            }
            if (plotArea.GetFirstChild<AreaChart>() is AreaChart areaChart) {
                GroupingValues grouping = areaChart.GetFirstChild<Grouping>()?.Val ?? GroupingValues.Standard;
                if (grouping == GroupingValues.PercentStacked) return OfficeChartType.AreaStacked100;
                return grouping == GroupingValues.Stacked ? OfficeChartType.AreaStacked : OfficeChartType.Area;
            }
            if (plotArea.GetFirstChild<OfPieChart>() is OfPieChart ofPieChart) {
                OfPieValues type = ofPieChart.GetFirstChild<OfPieType>()?.Val ?? OfPieValues.Pie;
                return type == OfPieValues.Bar ? OfficeChartType.BarOfPie : OfficeChartType.PieOfPie;
            }
            if (plotArea.GetFirstChild<Pie3DChart>() != null) return OfficeChartType.Pie3D;
            if (plotArea.GetFirstChild<PieChart>() != null) return OfficeChartType.Pie;
            if (plotArea.GetFirstChild<DoughnutChart>() != null) return OfficeChartType.Doughnut;
            if (plotArea.GetFirstChild<ScatterChart>() != null) return OfficeChartType.Scatter;
            if (plotArea.GetFirstChild<BubbleChart>() != null) return OfficeChartType.Bubble;
            if (plotArea.GetFirstChild<RadarChart>() != null) return OfficeChartType.Radar;
            if (plotArea.GetFirstChild<StockChart>() != null) return OfficeChartType.Stock;
            if (plotArea.GetFirstChild<Surface3DChart>() is Surface3DChart surface3DChart) {
                bool wireframe = surface3DChart.GetFirstChild<Wireframe>()?.Val ?? false;
                return wireframe ? OfficeChartType.SurfaceWireframe : OfficeChartType.Surface;
            }
            if (plotArea.GetFirstChild<SurfaceChart>() is SurfaceChart surfaceChart) {
                bool wireframe = surfaceChart.GetFirstChild<Wireframe>()?.Val ?? false;
                return wireframe ? OfficeChartType.SurfaceContourWireframe : OfficeChartType.SurfaceContour;
            }
            return OfficeChartType.ColumnClustered;
        }

        private static bool IsBarChartType(OfficeChartType chartType) {
            return chartType == OfficeChartType.BarClustered
                || chartType == OfficeChartType.BarStacked
                || chartType == OfficeChartType.BarStacked100
                || chartType == OfficeChartType.Bar3DClustered
                || chartType == OfficeChartType.Bar3DStacked
                || chartType == OfficeChartType.Bar3DStacked100;
        }

        private static bool IsBar3DChartType(OfficeChartType chartType) {
            return chartType == OfficeChartType.Column3DClustered
                || chartType == OfficeChartType.Column3DStacked
                || chartType == OfficeChartType.Column3DStacked100
                || chartType == OfficeChartType.Bar3DClustered
                || chartType == OfficeChartType.Bar3DStacked
                || chartType == OfficeChartType.Bar3DStacked100;
        }

        private static bool IsArea3DChartType(OfficeChartType chartType) {
            return chartType == OfficeChartType.Area3D
                || chartType == OfficeChartType.Area3DStacked
                || chartType == OfficeChartType.Area3DStacked100;
        }

        private static bool IsSingleSeriesPieVariant(OfficeChartType chartType) {
            return chartType == OfficeChartType.Pie3D
                || chartType == OfficeChartType.PieOfPie
                || chartType == OfficeChartType.BarOfPie;
        }

        private static void ValidateSingleSeriesPieVariants(IReadOnlyList<SeriesDescriptor> descriptors) {
            foreach (var group in descriptors.Where(descriptor => IsSingleSeriesPieVariant(descriptor.ChartType)).GroupBy(descriptor => descriptor.ChartType)) {
                if (group.Count() <= 1) continue;
                string chartName = group.Key switch {
                    OfficeChartType.Pie3D => "3-D pie",
                    OfficeChartType.PieOfPie => "pie-of-pie",
                    OfficeChartType.BarOfPie => "bar-of-pie",
                    _ => "pie"
                };
                throw new NotSupportedException($"{chartName} charts support exactly one series.");
            }
        }

        private static bool IsSurfaceChartType(OfficeChartType chartType) {
            return chartType == OfficeChartType.Surface
                || chartType == OfficeChartType.SurfaceWireframe
                || chartType == OfficeChartType.SurfaceContour
                || chartType == OfficeChartType.SurfaceContourWireframe;
        }

        private static SeriesDescriptorSummary SummarizeSeriesDescriptors(IReadOnlyList<SeriesDescriptor> descriptors) {
            bool hasSecondary = false;
            bool hasScatter = false;
            bool hasBubble = false;
            bool hasStock = false;
            bool hasSurface = false;
            bool hasLine3D = false;
            bool hasBar3D = false;
            bool hasArea3D = false;
            bool hasBar = false;
            bool hasNonBar = false;
            bool hasPieOrDoughnut = false;
            bool hasMultipleTypes = false;
            bool hasFirstType = false;
            OfficeChartType firstType = default;

            for (int i = 0; i < descriptors.Count; i++) {
                SeriesDescriptor descriptor = descriptors[i];
                OfficeChartType chartType = descriptor.ChartType;
                if (!hasFirstType) {
                    firstType = chartType;
                    hasFirstType = true;
                } else if (chartType != firstType) {
                    hasMultipleTypes = true;
                }

                if (descriptor.AxisGroup == OfficeChartAxisGroup.Secondary) {
                    hasSecondary = true;
                }

                if (chartType == OfficeChartType.Scatter) {
                    hasScatter = true;
                } else if (chartType == OfficeChartType.Bubble) {
                    hasBubble = true;
                } else if (chartType == OfficeChartType.Stock) {
                    hasStock = true;
                } else if (IsSurfaceChartType(chartType)) {
                    hasSurface = true;
                } else if (chartType == OfficeChartType.Line3D) {
                    hasLine3D = true;
                } else if (IsBar3DChartType(chartType)) {
                    hasBar3D = true;
                } else if (IsArea3DChartType(chartType)) {
                    hasArea3D = true;
                } else if (chartType == OfficeChartType.Pie
                    || chartType == OfficeChartType.Pie3D
                    || chartType == OfficeChartType.PieOfPie
                    || chartType == OfficeChartType.BarOfPie
                    || chartType == OfficeChartType.Doughnut) {
                    hasPieOrDoughnut = true;
                }

                if (IsBarChartType(chartType)) {
                    hasBar = true;
                } else {
                    hasNonBar = true;
                }
            }

            return new SeriesDescriptorSummary(hasSecondary, hasScatter, hasBubble, hasStock, hasSurface, hasLine3D, hasBar3D, hasArea3D, hasBar, hasNonBar, hasPieOrDoughnut, hasMultipleTypes);
        }

        private static List<SeriesDescriptorGroup> GroupSeriesDescriptors(IReadOnlyList<SeriesDescriptor> descriptors) {
            var groups = new List<SeriesDescriptorGroup>();
            for (int i = 0; i < descriptors.Count; i++) {
                SeriesDescriptor descriptor = descriptors[i];
                SeriesDescriptorGroup? group = null;
                for (int g = 0; g < groups.Count; g++) {
                    SeriesDescriptorGroup candidate = groups[g];
                    if (candidate.ChartType == descriptor.ChartType && candidate.AxisGroup == descriptor.AxisGroup) {
                        group = candidate;
                        break;
                    }
                }

                if (group == null) {
                    group = new SeriesDescriptorGroup(descriptor.ChartType, descriptor.AxisGroup);
                    groups.Add(group);
                }

                group.Descriptors.Add(descriptor);
            }

            return groups;
        }

        private static GroupingValues GetAreaGrouping(OfficeChartType chartType) {
            return chartType switch {
                OfficeChartType.AreaStacked => GroupingValues.Stacked,
                OfficeChartType.AreaStacked100 => GroupingValues.PercentStacked,
                OfficeChartType.Area3DStacked => GroupingValues.Stacked,
                OfficeChartType.Area3DStacked100 => GroupingValues.PercentStacked,
                _ => GroupingValues.Standard
            };
        }

        private static GroupingValues GetLineGrouping(OfficeChartType chartType) {
            return chartType switch {
                OfficeChartType.LineStacked => GroupingValues.Stacked,
                OfficeChartType.LineStacked100 => GroupingValues.PercentStacked,
                _ => GroupingValues.Standard
            };
        }

        private static (BarDirectionValues Direction, BarGroupingValues Grouping) GetBarChartSettings(OfficeChartType chartType) {
            return chartType switch {
                OfficeChartType.ColumnClustered => (BarDirectionValues.Column, BarGroupingValues.Clustered),
                OfficeChartType.ColumnStacked => (BarDirectionValues.Column, BarGroupingValues.Stacked),
                OfficeChartType.ColumnStacked100 => (BarDirectionValues.Column, BarGroupingValues.PercentStacked),
                OfficeChartType.Column3DClustered => (BarDirectionValues.Column, BarGroupingValues.Clustered),
                OfficeChartType.Column3DStacked => (BarDirectionValues.Column, BarGroupingValues.Stacked),
                OfficeChartType.Column3DStacked100 => (BarDirectionValues.Column, BarGroupingValues.PercentStacked),
                OfficeChartType.BarClustered => (BarDirectionValues.Bar, BarGroupingValues.Clustered),
                OfficeChartType.BarStacked => (BarDirectionValues.Bar, BarGroupingValues.Stacked),
                OfficeChartType.BarStacked100 => (BarDirectionValues.Bar, BarGroupingValues.PercentStacked),
                OfficeChartType.Bar3DClustered => (BarDirectionValues.Bar, BarGroupingValues.Clustered),
                OfficeChartType.Bar3DStacked => (BarDirectionValues.Bar, BarGroupingValues.Stacked),
                OfficeChartType.Bar3DStacked100 => (BarDirectionValues.Bar, BarGroupingValues.PercentStacked),
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
