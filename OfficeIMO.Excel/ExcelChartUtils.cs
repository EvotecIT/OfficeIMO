using System.Globalization;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    internal static partial class ExcelChartUtils {
        private const string ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        private const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        // Schema tokens of the chart enumerations this code reads and writes.
        private const string BarDirectionBar = "bar";
        private const string BarDirectionColumn = "col";
        private const string BarGroupingClustered = "clustered";
        private const string GroupingStandard = "standard";
        private const string GroupingStacked = "stacked";
        private const string GroupingPercentStacked = "percentStacked";
        private const string OfPieTypePie = "pie";
        private const string OfPieTypeBar = "bar";
        private const string AxisPositionBottom = "b";
        private const string AxisPositionLeft = "l";
        private const string AxisPositionRight = "r";
        private const string AxisPositionTop = "t";

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
            ArgumentException.ThrowIfNullOrWhiteSpace(a1Range);

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

        internal static OfficeChartType InferChartType(OfficeXmlNode plotArea) {
            OfficeXmlNode barChart = plotArea.Child(C, "barChart");
            if (barChart.Exists) {
                string direction = barChart.ChildVal(C, "barDir") ?? BarDirectionColumn;
                string grouping = barChart.ChildVal(C, "grouping") ?? BarGroupingClustered;
                if (direction == BarDirectionBar) {
                    if (grouping == GroupingPercentStacked) return OfficeChartType.BarStacked100;
                    return grouping == GroupingStacked ? OfficeChartType.BarStacked : OfficeChartType.BarClustered;
                }
                if (grouping == GroupingPercentStacked) return OfficeChartType.ColumnStacked100;
                return grouping == GroupingStacked ? OfficeChartType.ColumnStacked : OfficeChartType.ColumnClustered;
            }
            OfficeXmlNode bar3DChart = plotArea.Child(C, "bar3DChart");
            if (bar3DChart.Exists) {
                string direction = bar3DChart.ChildVal(C, "barDir") ?? BarDirectionColumn;
                string grouping = bar3DChart.ChildVal(C, "grouping") ?? BarGroupingClustered;
                if (direction == BarDirectionBar) {
                    if (grouping == GroupingPercentStacked) return OfficeChartType.Bar3DStacked100;
                    return grouping == GroupingStacked ? OfficeChartType.Bar3DStacked : OfficeChartType.Bar3DClustered;
                }
                if (grouping == GroupingPercentStacked) return OfficeChartType.Column3DStacked100;
                return grouping == GroupingStacked ? OfficeChartType.Column3DStacked : OfficeChartType.Column3DClustered;
            }
            if (plotArea.Child(C, "line3DChart").Exists) return OfficeChartType.Line3D;
            OfficeXmlNode lineChart = plotArea.Child(C, "lineChart");
            if (lineChart.Exists) {
                string grouping = lineChart.ChildVal(C, "grouping") ?? GroupingStandard;
                if (grouping == GroupingPercentStacked) return OfficeChartType.LineStacked100;
                return grouping == GroupingStacked ? OfficeChartType.LineStacked : OfficeChartType.Line;
            }
            OfficeXmlNode area3DChart = plotArea.Child(C, "area3DChart");
            if (area3DChart.Exists) {
                string grouping = area3DChart.ChildVal(C, "grouping") ?? GroupingStandard;
                if (grouping == GroupingPercentStacked) return OfficeChartType.Area3DStacked100;
                return grouping == GroupingStacked ? OfficeChartType.Area3DStacked : OfficeChartType.Area3D;
            }
            OfficeXmlNode areaChart = plotArea.Child(C, "areaChart");
            if (areaChart.Exists) {
                string grouping = areaChart.ChildVal(C, "grouping") ?? GroupingStandard;
                if (grouping == GroupingPercentStacked) return OfficeChartType.AreaStacked100;
                return grouping == GroupingStacked ? OfficeChartType.AreaStacked : OfficeChartType.Area;
            }
            OfficeXmlNode ofPieChart = plotArea.Child(C, "ofPieChart");
            if (ofPieChart.Exists) {
                string type = ofPieChart.ChildVal(C, "ofPieType") ?? OfPieTypePie;
                return type == OfPieTypeBar ? OfficeChartType.BarOfPie : OfficeChartType.PieOfPie;
            }
            if (plotArea.Child(C, "pie3DChart").Exists) return OfficeChartType.Pie3D;
            if (plotArea.Child(C, "pieChart").Exists) return OfficeChartType.Pie;
            if (plotArea.Child(C, "doughnutChart").Exists) return OfficeChartType.Doughnut;
            if (plotArea.Child(C, "scatterChart").Exists) return OfficeChartType.Scatter;
            if (plotArea.Child(C, "bubbleChart").Exists) return OfficeChartType.Bubble;
            if (plotArea.Child(C, "radarChart").Exists) return OfficeChartType.Radar;
            if (plotArea.Child(C, "stockChart").Exists) return OfficeChartType.Stock;
            OfficeXmlNode surface3DChart = plotArea.Child(C, "surface3DChart");
            if (surface3DChart.Exists) {
                bool wireframe = surface3DChart.ChildBoolVal(C, "wireframe") ?? false;
                return wireframe ? OfficeChartType.SurfaceWireframe : OfficeChartType.Surface;
            }
            OfficeXmlNode surfaceChart = plotArea.Child(C, "surfaceChart");
            if (surfaceChart.Exists) {
                bool wireframe = surfaceChart.ChildBoolVal(C, "wireframe") ?? false;
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

        private static string GetAreaGrouping(OfficeChartType chartType) {
            return chartType switch {
                OfficeChartType.AreaStacked => GroupingStacked,
                OfficeChartType.AreaStacked100 => GroupingPercentStacked,
                OfficeChartType.Area3DStacked => GroupingStacked,
                OfficeChartType.Area3DStacked100 => GroupingPercentStacked,
                _ => GroupingStandard
            };
        }

        private static string GetLineGrouping(OfficeChartType chartType) {
            return chartType switch {
                OfficeChartType.LineStacked => GroupingStacked,
                OfficeChartType.LineStacked100 => GroupingPercentStacked,
                _ => GroupingStandard
            };
        }

        private static (string Direction, string Grouping) GetBarChartSettings(OfficeChartType chartType) {
            return chartType switch {
                OfficeChartType.ColumnClustered => (BarDirectionColumn, BarGroupingClustered),
                OfficeChartType.ColumnStacked => (BarDirectionColumn, GroupingStacked),
                OfficeChartType.ColumnStacked100 => (BarDirectionColumn, GroupingPercentStacked),
                OfficeChartType.Column3DClustered => (BarDirectionColumn, BarGroupingClustered),
                OfficeChartType.Column3DStacked => (BarDirectionColumn, GroupingStacked),
                OfficeChartType.Column3DStacked100 => (BarDirectionColumn, GroupingPercentStacked),
                OfficeChartType.BarClustered => (BarDirectionBar, BarGroupingClustered),
                OfficeChartType.BarStacked => (BarDirectionBar, GroupingStacked),
                OfficeChartType.BarStacked100 => (BarDirectionBar, GroupingPercentStacked),
                OfficeChartType.Bar3DClustered => (BarDirectionBar, BarGroupingClustered),
                OfficeChartType.Bar3DStacked => (BarDirectionBar, GroupingStacked),
                OfficeChartType.Bar3DStacked100 => (BarDirectionBar, GroupingPercentStacked),
                _ => (BarDirectionColumn, BarGroupingClustered)
            };
        }

        private static IReadOnlyList<double> ParseNumericCategories(IReadOnlyList<string> categories) {
            var values = new double[categories.Count];
            for (int i = 0; i < categories.Count; i++) {
                if (!OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(categories[i], NumberStyles.Any, CultureInfo.InvariantCulture, out double val)) {
                    val = 0d;
                }
                values[i] = val;
            }
            return values;
        }
    }
}
