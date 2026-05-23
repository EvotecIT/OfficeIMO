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
            public SeriesDescriptorGroup(ExcelChartType chartType, ExcelChartAxisGroup axisGroup) {
                ChartType = chartType;
                AxisGroup = axisGroup;
                Descriptors = new List<SeriesDescriptor>();
            }

            public ExcelChartType ChartType { get; }
            public ExcelChartAxisGroup AxisGroup { get; }
            public List<SeriesDescriptor> Descriptors { get; }
        }

        internal static string BuildCellA1(int row, int column) {
            return A1.AbsoluteCellReference(row, column);
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
                return A1.TryParseRange(normalized, out r1, out c1, out r2, out c2);
            }
            var (row, col) = A1.ParseCellRef(normalized);
            if (row <= 0 || col <= 0) return false;
            r1 = r2 = row;
            c1 = c2 = col;
            return true;
        }

        private static List<SeriesDescriptor> BuildSeriesDescriptors(ExcelChartDataRange range, ExcelChartData? data,
            ExcelChartType defaultType, bool useSeriesOverrides = true) {
            IReadOnlyList<ExcelChartSeries>? seriesList = data?.Series;
            int count = seriesList?.Count ?? range.SeriesCount;
            var descriptors = new List<SeriesDescriptor>(count);
            for (int i = 0; i < count; i++) {
                ExcelChartSeries? series = seriesList != null && i < seriesList.Count ? seriesList[i] : null;
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
                    if (grouping == BarGroupingValues.PercentStacked) return ExcelChartType.BarStacked100;
                    return grouping == BarGroupingValues.Stacked ? ExcelChartType.BarStacked : ExcelChartType.BarClustered;
                }
                if (grouping == BarGroupingValues.PercentStacked) return ExcelChartType.ColumnStacked100;
                return grouping == BarGroupingValues.Stacked ? ExcelChartType.ColumnStacked : ExcelChartType.ColumnClustered;
            }
            if (plotArea.GetFirstChild<Bar3DChart>() is Bar3DChart bar3DChart) {
                BarDirectionValues direction = bar3DChart.GetFirstChild<BarDirection>()?.Val ?? BarDirectionValues.Column;
                BarGroupingValues grouping = bar3DChart.GetFirstChild<BarGrouping>()?.Val ?? BarGroupingValues.Clustered;
                if (direction == BarDirectionValues.Bar) {
                    if (grouping == BarGroupingValues.PercentStacked) return ExcelChartType.Bar3DStacked100;
                    return grouping == BarGroupingValues.Stacked ? ExcelChartType.Bar3DStacked : ExcelChartType.Bar3DClustered;
                }
                if (grouping == BarGroupingValues.PercentStacked) return ExcelChartType.Column3DStacked100;
                return grouping == BarGroupingValues.Stacked ? ExcelChartType.Column3DStacked : ExcelChartType.Column3DClustered;
            }
            if (plotArea.GetFirstChild<Line3DChart>() != null) return ExcelChartType.Line3D;
            if (plotArea.GetFirstChild<LineChart>() is LineChart lineChart) {
                GroupingValues grouping = lineChart.GetFirstChild<Grouping>()?.Val ?? GroupingValues.Standard;
                if (grouping == GroupingValues.PercentStacked) return ExcelChartType.LineStacked100;
                return grouping == GroupingValues.Stacked ? ExcelChartType.LineStacked : ExcelChartType.Line;
            }
            if (plotArea.GetFirstChild<Area3DChart>() is Area3DChart area3DChart) {
                GroupingValues grouping = area3DChart.GetFirstChild<Grouping>()?.Val ?? GroupingValues.Standard;
                if (grouping == GroupingValues.PercentStacked) return ExcelChartType.Area3DStacked100;
                return grouping == GroupingValues.Stacked ? ExcelChartType.Area3DStacked : ExcelChartType.Area3D;
            }
            if (plotArea.GetFirstChild<AreaChart>() is AreaChart areaChart) {
                GroupingValues grouping = areaChart.GetFirstChild<Grouping>()?.Val ?? GroupingValues.Standard;
                if (grouping == GroupingValues.PercentStacked) return ExcelChartType.AreaStacked100;
                return grouping == GroupingValues.Stacked ? ExcelChartType.AreaStacked : ExcelChartType.Area;
            }
            if (plotArea.GetFirstChild<OfPieChart>() is OfPieChart ofPieChart) {
                OfPieValues type = ofPieChart.GetFirstChild<OfPieType>()?.Val ?? OfPieValues.Pie;
                return type == OfPieValues.Bar ? ExcelChartType.BarOfPie : ExcelChartType.PieOfPie;
            }
            if (plotArea.GetFirstChild<Pie3DChart>() != null) return ExcelChartType.Pie3D;
            if (plotArea.GetFirstChild<PieChart>() != null) return ExcelChartType.Pie;
            if (plotArea.GetFirstChild<DoughnutChart>() != null) return ExcelChartType.Doughnut;
            if (plotArea.GetFirstChild<ScatterChart>() != null) return ExcelChartType.Scatter;
            if (plotArea.GetFirstChild<BubbleChart>() != null) return ExcelChartType.Bubble;
            if (plotArea.GetFirstChild<RadarChart>() != null) return ExcelChartType.Radar;
            if (plotArea.GetFirstChild<StockChart>() != null) return ExcelChartType.Stock;
            if (plotArea.GetFirstChild<Surface3DChart>() is Surface3DChart surface3DChart) {
                bool wireframe = surface3DChart.GetFirstChild<Wireframe>()?.Val ?? false;
                return wireframe ? ExcelChartType.SurfaceWireframe : ExcelChartType.Surface;
            }
            if (plotArea.GetFirstChild<SurfaceChart>() is SurfaceChart surfaceChart) {
                bool wireframe = surfaceChart.GetFirstChild<Wireframe>()?.Val ?? false;
                return wireframe ? ExcelChartType.SurfaceContourWireframe : ExcelChartType.SurfaceContour;
            }
            return ExcelChartType.ColumnClustered;
        }

        private static bool IsBarChartType(ExcelChartType chartType) {
            return chartType == ExcelChartType.BarClustered
                || chartType == ExcelChartType.BarStacked
                || chartType == ExcelChartType.BarStacked100
                || chartType == ExcelChartType.Bar3DClustered
                || chartType == ExcelChartType.Bar3DStacked
                || chartType == ExcelChartType.Bar3DStacked100;
        }

        private static bool IsBar3DChartType(ExcelChartType chartType) {
            return chartType == ExcelChartType.Column3DClustered
                || chartType == ExcelChartType.Column3DStacked
                || chartType == ExcelChartType.Column3DStacked100
                || chartType == ExcelChartType.Bar3DClustered
                || chartType == ExcelChartType.Bar3DStacked
                || chartType == ExcelChartType.Bar3DStacked100;
        }

        private static bool IsArea3DChartType(ExcelChartType chartType) {
            return chartType == ExcelChartType.Area3D
                || chartType == ExcelChartType.Area3DStacked
                || chartType == ExcelChartType.Area3DStacked100;
        }

        private static bool IsSurfaceChartType(ExcelChartType chartType) {
            return chartType == ExcelChartType.Surface
                || chartType == ExcelChartType.SurfaceWireframe
                || chartType == ExcelChartType.SurfaceContour
                || chartType == ExcelChartType.SurfaceContourWireframe;
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
            ExcelChartType firstType = default;

            for (int i = 0; i < descriptors.Count; i++) {
                SeriesDescriptor descriptor = descriptors[i];
                ExcelChartType chartType = descriptor.ChartType;
                if (!hasFirstType) {
                    firstType = chartType;
                    hasFirstType = true;
                } else if (chartType != firstType) {
                    hasMultipleTypes = true;
                }

                if (descriptor.AxisGroup == ExcelChartAxisGroup.Secondary) {
                    hasSecondary = true;
                }

                if (chartType == ExcelChartType.Scatter) {
                    hasScatter = true;
                } else if (chartType == ExcelChartType.Bubble) {
                    hasBubble = true;
                } else if (chartType == ExcelChartType.Stock) {
                    hasStock = true;
                } else if (IsSurfaceChartType(chartType)) {
                    hasSurface = true;
                } else if (chartType == ExcelChartType.Line3D) {
                    hasLine3D = true;
                } else if (IsBar3DChartType(chartType)) {
                    hasBar3D = true;
                } else if (IsArea3DChartType(chartType)) {
                    hasArea3D = true;
                } else if (chartType == ExcelChartType.Pie
                    || chartType == ExcelChartType.Pie3D
                    || chartType == ExcelChartType.PieOfPie
                    || chartType == ExcelChartType.BarOfPie
                    || chartType == ExcelChartType.Doughnut) {
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

        private static GroupingValues GetAreaGrouping(ExcelChartType chartType) {
            return chartType switch {
                ExcelChartType.AreaStacked => GroupingValues.Stacked,
                ExcelChartType.AreaStacked100 => GroupingValues.PercentStacked,
                ExcelChartType.Area3DStacked => GroupingValues.Stacked,
                ExcelChartType.Area3DStacked100 => GroupingValues.PercentStacked,
                _ => GroupingValues.Standard
            };
        }

        private static GroupingValues GetLineGrouping(ExcelChartType chartType) {
            return chartType switch {
                ExcelChartType.LineStacked => GroupingValues.Stacked,
                ExcelChartType.LineStacked100 => GroupingValues.PercentStacked,
                _ => GroupingValues.Standard
            };
        }

        private static (BarDirectionValues Direction, BarGroupingValues Grouping) GetBarChartSettings(ExcelChartType chartType) {
            return chartType switch {
                ExcelChartType.ColumnClustered => (BarDirectionValues.Column, BarGroupingValues.Clustered),
                ExcelChartType.ColumnStacked => (BarDirectionValues.Column, BarGroupingValues.Stacked),
                ExcelChartType.ColumnStacked100 => (BarDirectionValues.Column, BarGroupingValues.PercentStacked),
                ExcelChartType.Column3DClustered => (BarDirectionValues.Column, BarGroupingValues.Clustered),
                ExcelChartType.Column3DStacked => (BarDirectionValues.Column, BarGroupingValues.Stacked),
                ExcelChartType.Column3DStacked100 => (BarDirectionValues.Column, BarGroupingValues.PercentStacked),
                ExcelChartType.BarClustered => (BarDirectionValues.Bar, BarGroupingValues.Clustered),
                ExcelChartType.BarStacked => (BarDirectionValues.Bar, BarGroupingValues.Stacked),
                ExcelChartType.BarStacked100 => (BarDirectionValues.Bar, BarGroupingValues.PercentStacked),
                ExcelChartType.Bar3DClustered => (BarDirectionValues.Bar, BarGroupingValues.Clustered),
                ExcelChartType.Bar3DStacked => (BarDirectionValues.Bar, BarGroupingValues.Stacked),
                ExcelChartType.Bar3DStacked100 => (BarDirectionValues.Bar, BarGroupingValues.PercentStacked),
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
