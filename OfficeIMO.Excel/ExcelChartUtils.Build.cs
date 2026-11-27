using System.Text;
using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    internal static partial class ExcelChartUtils {
        // The chart space root every built chart starts from, with the namespaces declared in this order.
        private static readonly byte[] ChartSpaceRootMarkup = Encoding.UTF8.GetBytes(
            "<c:chartSpace xmlns:c=\"" + ChartNamespace + "\" xmlns:a=\"" + DrawingNamespace + "\" xmlns:r=\"" + RelationshipNamespace + "\" />");

        /// <summary>
        /// Builds a chart space over a new chart part: the root first, as the context the children are made in, and the
        /// children through <paramref name="build"/>. A build that fails leaves the part as empty as it was.
        /// </summary>
        private static void BuildChartSpace(ExcelPartRef chartPart, Action<OfficeXmlNode> build) {
            OfficeXmlNode chartSpace = chartPart.SetRootMarkup(ChartSpaceRootMarkup);
            try {
                build(chartSpace);
            } catch {
                chartPart.Feed(Array.Empty<byte>());
                throw;
            }
        }

        private static OfficeXmlNode CreateChartElement(OfficeXmlNode chartSpace, string? title) {
            chartSpace.Append(chartSpace.NewVal(C, "date1904", false));
            chartSpace.Append(chartSpace.NewVal(C, "lang", "en-US"));

            OfficeXmlNode chart = chartSpace.New(C, "chart");
            if (!string.IsNullOrWhiteSpace(title)) {
                chart.Append(CreateChartTitle(chart, title!));
            }
            chart.Append(chart.NewVal(C, "autoTitleDeleted", string.IsNullOrWhiteSpace(title)));
            return chart;
        }

        private static OfficeXmlNode CreatePlotArea(OfficeXmlNode chart) {
            OfficeXmlNode plotArea = chart.New(C, "plotArea");
            plotArea.Append(plotArea.New(C, "layout"));
            return plotArea;
        }

        private static void AppendChartTail(OfficeXmlNode chart) {
            OfficeXmlNode legend = chart.New(C, "legend");
            legend.Append(legend.NewVal(C, "legendPos", "b"));
            legend.Append(legend.New(C, "layout"));
            legend.Append(legend.NewVal(C, "overlay", false));
            chart.Append(legend);
            chart.Append(chart.NewVal(C, "plotVisOnly", true));
            chart.Append(chart.NewVal(C, "dispBlanksAs", "gap"));
            chart.Append(chart.NewVal(C, "showDLblsOverMax", false));
        }

        internal static void PopulateChart(ExcelPartRef chartPart, OfficeChartType type, ExcelChartDataRange range,
            ExcelChartData? data, string? title = null, IReadOnlyList<string>? nativeSeriesNames = null) {
            if (chartPart.IsNone) throw new ArgumentNullException(nameof(chartPart));
            BuildChartSpace(chartPart, chartSpace => PopulateChartSpace(chartSpace, type, range, data, title, nativeSeriesNames));
        }

        private static void PopulateChartSpace(OfficeXmlNode chartSpace, OfficeChartType type, ExcelChartDataRange range,
            ExcelChartData? data, string? title, IReadOnlyList<string>? nativeSeriesNames) {
            OfficeXmlNode chart = CreateChartElement(chartSpace, title);

            OfficeXmlNode plotArea = CreatePlotArea(chart);
            List<SeriesDescriptor> descriptors = BuildSeriesDescriptors(range, data, type);
            SeriesDescriptorSummary summary = SummarizeSeriesDescriptors(descriptors);
            if (summary.HasPieOrDoughnut) {
                ValidateSingleSeriesPieVariants(descriptors);
            }
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
                plotArea.Append(CreateScatterChart(plotArea, range, descriptors, xAxisId, yAxisId, data));
                plotArea.Append(CreateValueAxis(plotArea, xAxisId, yAxisId, AxisPositionBottom));
                plotArea.Append(CreateValueAxis(plotArea, yAxisId, xAxisId, AxisPositionLeft));
            } else if (summary.HasMultipleTypes || summary.HasSecondary) {
                BuildComboPlotArea(plotArea, range, descriptors, summary);
            } else {
                OfficeChartType chartType = descriptors.Count > 0 ? descriptors[0].ChartType : type;
                uint categoryAxisId = ExcelChartAxisIdGenerator.GetNextId();
                uint valueAxisId = ExcelChartAxisIdGenerator.GetNextId();
                uint seriesAxisId = IsSurfaceChartType(chartType) || chartType == OfficeChartType.Line3D || IsBar3DChartType(chartType) || IsArea3DChartType(chartType) ? ExcelChartAxisIdGenerator.GetNextId() : 0;

                switch (chartType) {
                    case OfficeChartType.ColumnClustered:
                        plotArea.Append(CreateBarChart(plotArea, range, descriptors, BarDirectionColumn, BarGroupingClustered, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId));
                        break;
                    case OfficeChartType.ColumnStacked:
                        plotArea.Append(CreateBarChart(plotArea, range, descriptors, BarDirectionColumn, GroupingStacked, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId));
                        break;
                    case OfficeChartType.ColumnStacked100:
                        plotArea.Append(CreateBarChart(plotArea, range, descriptors, BarDirectionColumn, GroupingPercentStacked, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId));
                        break;
                    case OfficeChartType.BarClustered:
                        plotArea.Append(CreateBarChart(plotArea, range, descriptors, BarDirectionBar, BarGroupingClustered, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId, AxisPositionLeft));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId, AxisPositionBottom));
                        break;
                    case OfficeChartType.BarStacked:
                        plotArea.Append(CreateBarChart(plotArea, range, descriptors, BarDirectionBar, GroupingStacked, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId, AxisPositionLeft));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId, AxisPositionBottom));
                        break;
                    case OfficeChartType.BarStacked100:
                        plotArea.Append(CreateBarChart(plotArea, range, descriptors, BarDirectionBar, GroupingPercentStacked, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId, AxisPositionLeft));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId, AxisPositionBottom));
                        break;
                    case OfficeChartType.Column3DClustered:
                        plotArea.Append(CreateBar3DChart(plotArea, range, descriptors, BarDirectionColumn, BarGroupingClustered, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId));
                        plotArea.Append(CreateSeriesAxis(plotArea, seriesAxisId, valueAxisId));
                        break;
                    case OfficeChartType.Column3DStacked:
                        plotArea.Append(CreateBar3DChart(plotArea, range, descriptors, BarDirectionColumn, GroupingStacked, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId));
                        plotArea.Append(CreateSeriesAxis(plotArea, seriesAxisId, valueAxisId));
                        break;
                    case OfficeChartType.Column3DStacked100:
                        plotArea.Append(CreateBar3DChart(plotArea, range, descriptors, BarDirectionColumn, GroupingPercentStacked, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId));
                        plotArea.Append(CreateSeriesAxis(plotArea, seriesAxisId, valueAxisId));
                        break;
                    case OfficeChartType.Bar3DClustered:
                        plotArea.Append(CreateBar3DChart(plotArea, range, descriptors, BarDirectionBar, BarGroupingClustered, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId, AxisPositionLeft));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId, AxisPositionBottom));
                        plotArea.Append(CreateSeriesAxis(plotArea, seriesAxisId, valueAxisId));
                        break;
                    case OfficeChartType.Bar3DStacked:
                        plotArea.Append(CreateBar3DChart(plotArea, range, descriptors, BarDirectionBar, GroupingStacked, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId, AxisPositionLeft));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId, AxisPositionBottom));
                        plotArea.Append(CreateSeriesAxis(plotArea, seriesAxisId, valueAxisId));
                        break;
                    case OfficeChartType.Bar3DStacked100:
                        plotArea.Append(CreateBar3DChart(plotArea, range, descriptors, BarDirectionBar, GroupingPercentStacked, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId, AxisPositionLeft));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId, AxisPositionBottom));
                        plotArea.Append(CreateSeriesAxis(plotArea, seriesAxisId, valueAxisId));
                        break;
                    case OfficeChartType.Line:
                        plotArea.Append(CreateLineChart(plotArea, range, descriptors, GroupingStandard, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId));
                        break;
                    case OfficeChartType.LineStacked:
                        plotArea.Append(CreateLineChart(plotArea, range, descriptors, GroupingStacked, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId));
                        break;
                    case OfficeChartType.LineStacked100:
                        plotArea.Append(CreateLineChart(plotArea, range, descriptors, GroupingPercentStacked, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId));
                        break;
                    case OfficeChartType.Line3D:
                        plotArea.Append(CreateLine3DChart(plotArea, range, descriptors, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId));
                        plotArea.Append(CreateSeriesAxis(plotArea, seriesAxisId, valueAxisId));
                        break;
                    case OfficeChartType.Area:
                        plotArea.Append(CreateAreaChart(plotArea, range, descriptors, GroupingStandard, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId));
                        break;
                    case OfficeChartType.AreaStacked:
                        plotArea.Append(CreateAreaChart(plotArea, range, descriptors, GroupingStacked, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId));
                        break;
                    case OfficeChartType.AreaStacked100:
                        plotArea.Append(CreateAreaChart(plotArea, range, descriptors, GroupingPercentStacked, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId));
                        break;
                    case OfficeChartType.Area3D:
                        plotArea.Append(CreateArea3DChart(plotArea, range, descriptors, GroupingStandard, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId));
                        plotArea.Append(CreateSeriesAxis(plotArea, seriesAxisId, valueAxisId));
                        break;
                    case OfficeChartType.Area3DStacked:
                        plotArea.Append(CreateArea3DChart(plotArea, range, descriptors, GroupingStacked, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId));
                        plotArea.Append(CreateSeriesAxis(plotArea, seriesAxisId, valueAxisId));
                        break;
                    case OfficeChartType.Area3DStacked100:
                        plotArea.Append(CreateArea3DChart(plotArea, range, descriptors, GroupingPercentStacked, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId));
                        plotArea.Append(CreateSeriesAxis(plotArea, seriesAxisId, valueAxisId));
                        break;
                    case OfficeChartType.Radar:
                        plotArea.Append(CreateRadarChart(plotArea, range, descriptors, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId));
                        break;
                    case OfficeChartType.Stock:
                        plotArea.Append(CreateStockChart(plotArea, range, descriptors, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId));
                        break;
                    case OfficeChartType.Surface:
                        plotArea.Append(CreateSurface3DChart(plotArea, range, descriptors, false, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId));
                        plotArea.Append(CreateSeriesAxis(plotArea, seriesAxisId, valueAxisId));
                        break;
                    case OfficeChartType.SurfaceWireframe:
                        plotArea.Append(CreateSurface3DChart(plotArea, range, descriptors, true, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId));
                        plotArea.Append(CreateSeriesAxis(plotArea, seriesAxisId, valueAxisId));
                        break;
                    case OfficeChartType.SurfaceContour:
                        plotArea.Append(CreateSurfaceChart(plotArea, range, descriptors, false, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId));
                        plotArea.Append(CreateSeriesAxis(plotArea, seriesAxisId, valueAxisId));
                        break;
                    case OfficeChartType.SurfaceContourWireframe:
                        plotArea.Append(CreateSurfaceChart(plotArea, range, descriptors, true, categoryAxisId, valueAxisId, seriesAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId));
                        plotArea.Append(CreateSeriesAxis(plotArea, seriesAxisId, valueAxisId));
                        break;
                    case OfficeChartType.Pie:
                        plotArea.Append(CreatePieChart(plotArea, range, descriptors));
                        break;
                    case OfficeChartType.Pie3D:
                        plotArea.Append(CreatePie3DChart(plotArea, range, descriptors));
                        break;
                    case OfficeChartType.PieOfPie:
                        plotArea.Append(CreateOfPieChart(plotArea, range, descriptors, OfPieTypePie));
                        break;
                    case OfficeChartType.BarOfPie:
                        plotArea.Append(CreateOfPieChart(plotArea, range, descriptors, OfPieTypeBar));
                        break;
                    case OfficeChartType.Doughnut:
                        plotArea.Append(CreateDoughnutChart(plotArea, range, descriptors));
                        break;
                    default:
                        plotArea.Append(CreateBarChart(plotArea, range, descriptors, BarDirectionColumn, BarGroupingClustered, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateCategoryAxis(plotArea, categoryAxisId, valueAxisId));
                        plotArea.Append(CreateValueAxis(plotArea, valueAxisId, categoryAxisId));
                        break;
                }
            }

            if (data != null) {
                foreach (OfficeXmlNode categories in plotArea.Descendants(C, "cat")) {
                    OfficeXmlNode reference = categories.Child(C, "strRef");
                    if (reference.Exists) {
                        reference.RemoveChildren(C, "strCache");
                        reference.Append(CreateStringCache(reference, data.Categories));
                    }
                }
            }

            chart.Append(plotArea);
            AppendChartTail(chart);

            // Table names remain available independently of optional numeric caches.
            // Series indexes identify the same columns across native chart families and groups.
            if (nativeSeriesNames != null) {
                foreach (OfficeXmlNode text in SeriesTexts(plotArea)) {
                    uint? index = text.Parent.Child(C, "idx").UInt32Val();
                    if (index.HasValue && index.Value < (uint)nativeSeriesNames.Count) {
                        string name = nativeSeriesNames[(int)index.Value];
                        OfficeXmlNode reference = text.Child(C, "strRef");
                        if (range.HasHeaderRow && reference.Exists) {
                            reference.RemoveChildren(C, "strCache");
                            reference.Append(CreateSingleStringCache(reference, name));
                        } else {
                            text.RemoveAllChildren();
                            text.Append(text.NewText(C, "v", name));
                        }
                    }
                }
            }
            chartSpace.Append(chart);
        }

        // The series name elements (c:tx directly under a c:ser) of a plot area, in document order.
        private static List<OfficeXmlNode> SeriesTexts(OfficeXmlNode plotArea) {
            var texts = new List<OfficeXmlNode>();
            foreach (OfficeXmlNode text in plotArea.Descendants(C, "tx")) {
                if (text.Parent.Is(C, "ser")) texts.Add(text);
            }
            return texts;
        }

        internal static void PopulateChartFromSeriesRanges(ExcelPartRef chartPart, OfficeChartType type, string defaultSheetName,
            IReadOnlyList<ExcelChartSeriesRange> seriesRanges, string? title = null) {
            if (chartPart.IsNone) throw new ArgumentNullException(nameof(chartPart));
            ArgumentException.ThrowIfNullOrWhiteSpace(defaultSheetName);
            ArgumentNullException.ThrowIfNull(seriesRanges);
            if (seriesRanges.Count == 0) {
                throw new ArgumentException("At least one series range is required.", nameof(seriesRanges));
            }
            if (type != OfficeChartType.Scatter && type != OfficeChartType.Bubble) {
                throw new NotSupportedException("Only scatter and bubble charts support explicit X/Y range definitions.");
            }
            if (type == OfficeChartType.Bubble && HasMissingBubbleSizeRange(seriesRanges)) {
                throw new ArgumentException("Bubble charts require bubble size ranges for each series.", nameof(seriesRanges));
            }

            BuildChartSpace(chartPart, chartSpace => {
            OfficeXmlNode chart = CreateChartElement(chartSpace, title);

            OfficeXmlNode plotArea = CreatePlotArea(chart);
            uint xAxisId = ExcelChartAxisIdGenerator.GetNextId();
            uint yAxisId = ExcelChartAxisIdGenerator.GetNextId();

            if (type == OfficeChartType.Scatter) {
                plotArea.Append(CreateScatterChartFromRanges(plotArea, seriesRanges, defaultSheetName, xAxisId, yAxisId));
            } else {
                plotArea.Append(CreateBubbleChartFromRanges(plotArea, seriesRanges, defaultSheetName, xAxisId, yAxisId));
            }

            plotArea.Append(CreateValueAxis(plotArea, xAxisId, yAxisId, AxisPositionBottom));
            plotArea.Append(CreateValueAxis(plotArea, yAxisId, xAxisId, AxisPositionLeft));

            chart.Append(plotArea);
            AppendChartTail(chart);

            chartSpace.Append(chart);
            });
        }

        private static void BuildComboPlotArea(OfficeXmlNode plotArea, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors, SeriesDescriptorSummary summary) {
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
            string primaryCategoryPosition = isBarOrientation ? AxisPositionLeft : AxisPositionBottom;
            string primaryValuePosition = isBarOrientation ? AxisPositionBottom : AxisPositionLeft;
            string secondaryCategoryPosition = isBarOrientation ? AxisPositionRight : AxisPositionTop;
            string secondaryValuePosition = isBarOrientation ? AxisPositionTop : AxisPositionRight;

            uint primaryCategoryId = ExcelChartAxisIdGenerator.GetNextId();
            uint primaryValueId = ExcelChartAxisIdGenerator.GetNextId();
            uint secondaryCategoryId = summary.HasSecondary ? ExcelChartAxisIdGenerator.GetNextId() : 0;
            uint secondaryValueId = summary.HasSecondary ? ExcelChartAxisIdGenerator.GetNextId() : 0;

            foreach (SeriesDescriptorGroup group in GroupSeriesDescriptors(descriptors)) {
                uint categoryAxisId = group.AxisGroup == OfficeChartAxisGroup.Secondary ? secondaryCategoryId : primaryCategoryId;
                uint valueAxisId = group.AxisGroup == OfficeChartAxisGroup.Secondary ? secondaryValueId : primaryValueId;

                switch (group.ChartType) {
                    case OfficeChartType.ColumnClustered:
                    case OfficeChartType.ColumnStacked:
                    case OfficeChartType.ColumnStacked100:
                    case OfficeChartType.BarClustered:
                    case OfficeChartType.BarStacked:
                    case OfficeChartType.BarStacked100: {
                        var settings = GetBarChartSettings(group.ChartType);
                        plotArea.Append(CreateBarChart(plotArea, range, group.Descriptors, settings.Direction, settings.Grouping, categoryAxisId, valueAxisId));
                        break;
                    }
                    case OfficeChartType.Line:
                    case OfficeChartType.LineStacked:
                    case OfficeChartType.LineStacked100:
                        plotArea.Append(CreateLineChart(plotArea, range, group.Descriptors, GetLineGrouping(group.ChartType), categoryAxisId, valueAxisId));
                        break;
                    case OfficeChartType.Area:
                    case OfficeChartType.AreaStacked:
                    case OfficeChartType.AreaStacked100:
                        plotArea.Append(CreateAreaChart(plotArea, range, group.Descriptors, GetAreaGrouping(group.ChartType), categoryAxisId, valueAxisId));
                        break;
                    default:
                        throw new NotSupportedException($"Chart type {group.ChartType} is not supported in combination charts.");
                }
            }

            plotArea.Append(CreateCategoryAxis(plotArea, primaryCategoryId, primaryValueId, primaryCategoryPosition));
            plotArea.Append(CreateValueAxis(plotArea, primaryValueId, primaryCategoryId, primaryValuePosition));
            if (summary.HasSecondary) {
                plotArea.Append(CreateCategoryAxis(plotArea, secondaryCategoryId, secondaryValueId, secondaryCategoryPosition));
                plotArea.Append(CreateValueAxis(plotArea, secondaryValueId, secondaryCategoryId, secondaryValuePosition));
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

        // Each builder below returns a detached element made with its intended parent as context; the caller inserts it.

        private static OfficeXmlNode NewAxisId(OfficeXmlNode parent, uint axisId) => parent.NewVal(C, "axId", axisId);

        private static OfficeXmlNode CreateBarChart(OfficeXmlNode parent, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors,
            string direction, string grouping, uint categoryAxisId, uint valueAxisId) {
            OfficeXmlNode barChart = parent.New(C, "barChart");
            barChart.Append(barChart.NewVal(C, "barDir", direction));
            barChart.Append(barChart.NewVal(C, "grouping", grouping));
            barChart.Append(barChart.NewVal(C, "varyColors", false));

            foreach (var descriptor in seriesDescriptors) {
                barChart.Append(CreateBarChartSeries(barChart, descriptor.Index, range, descriptor.Series));
            }

            barChart.Append(CreateDefaultDataLabels(barChart));
            barChart.Append(barChart.NewVal(C, "gapWidth", (uint)OfficeTheme.Default.ChartStyle.BarGapPercent));
            barChart.Append(barChart.NewVal(C, "overlap", -27));
            barChart.Append(NewAxisId(barChart, categoryAxisId));
            barChart.Append(NewAxisId(barChart, valueAxisId));
            return barChart;
        }

        // A new series with its index and order, as every builder starts one.
        private static OfficeXmlNode NewSeries(OfficeXmlNode chart, int seriesIndex) {
            OfficeXmlNode series = chart.New(C, "ser");
            series.Append(series.NewVal(C, "idx", (uint)seriesIndex));
            series.Append(series.NewVal(C, "order", (uint)seriesIndex));
            return series;
        }

        private static OfficeXmlNode CreateBarChartSeries(OfficeXmlNode chart, int seriesIndex, ExcelChartDataRange range, ExcelChartSeries? series) {
            OfficeXmlNode element = NewSeries(chart, seriesIndex);
            element.Append(CreateSeriesText(element, range, seriesIndex, series?.Name ?? $"Series {seriesIndex + 1}"));
            element.Append(element.NewVal(C, "invertIfNegative", false));
            element.Append(CreateCategoryAxisData(element, range));
            element.Append(CreateValues(element, range, seriesIndex, series));
            return element;
        }

        private static OfficeXmlNode CreateBar3DChart(OfficeXmlNode parent, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors,
            string direction, string grouping, uint categoryAxisId, uint valueAxisId, uint seriesAxisId) {
            OfficeXmlNode barChart = parent.New(C, "bar3DChart");
            barChart.Append(barChart.NewVal(C, "barDir", direction));
            barChart.Append(barChart.NewVal(C, "grouping", grouping));
            barChart.Append(barChart.NewVal(C, "varyColors", false));

            foreach (var descriptor in seriesDescriptors) {
                barChart.Append(CreateBarChartSeries(barChart, descriptor.Index, range, descriptor.Series));
            }

            barChart.Append(CreateDefaultDataLabels(barChart));
            barChart.Append(barChart.NewVal(C, "gapWidth", (uint)OfficeTheme.Default.ChartStyle.BarGapPercent));
            barChart.Append(barChart.NewVal(C, "gapDepth", 150U));
            barChart.Append(barChart.NewVal(C, "shape", "box"));
            barChart.Append(NewAxisId(barChart, categoryAxisId));
            barChart.Append(NewAxisId(barChart, valueAxisId));
            barChart.Append(NewAxisId(barChart, seriesAxisId));
            return barChart;
        }

        private static OfficeXmlNode CreateLineChart(OfficeXmlNode parent, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors,
            string grouping, uint categoryAxisId, uint valueAxisId) {
            OfficeXmlNode lineChart = parent.New(C, "lineChart");
            lineChart.Append(lineChart.NewVal(C, "grouping", grouping));
            lineChart.Append(lineChart.NewVal(C, "varyColors", false));

            foreach (var descriptor in seriesDescriptors) {
                lineChart.Append(CreateLineChartSeries(lineChart, descriptor.Index, range, descriptor.Series));
            }

            lineChart.Append(CreateDefaultDataLabels(lineChart));
            lineChart.Append(NewAxisId(lineChart, categoryAxisId));
            lineChart.Append(NewAxisId(lineChart, valueAxisId));
            return lineChart;
        }

        private static OfficeXmlNode CreateLineChartSeries(OfficeXmlNode chart, int seriesIndex, ExcelChartDataRange range, ExcelChartSeries? series) {
            OfficeXmlNode element = NewSeries(chart, seriesIndex);
            element.Append(CreateSeriesText(element, range, seriesIndex, series?.Name ?? $"Series {seriesIndex + 1}"));
            element.Append(CreateCategoryAxisData(element, range));
            element.Append(CreateValues(element, range, seriesIndex, series));
            element.Append(element.NewVal(C, "smooth", false));
            return element;
        }

        private static OfficeXmlNode CreateLine3DChart(OfficeXmlNode parent, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors,
            uint categoryAxisId, uint valueAxisId, uint seriesAxisId) {
            OfficeXmlNode lineChart = parent.New(C, "line3DChart");
            lineChart.Append(lineChart.NewVal(C, "grouping", GroupingStandard));
            lineChart.Append(lineChart.NewVal(C, "varyColors", false));

            foreach (var descriptor in seriesDescriptors) {
                lineChart.Append(CreateLineChartSeries(lineChart, descriptor.Index, range, descriptor.Series));
            }

            lineChart.Append(CreateDefaultDataLabels(lineChart));
            lineChart.Append(lineChart.NewVal(C, "gapDepth", 150U));
            lineChart.Append(NewAxisId(lineChart, categoryAxisId));
            lineChart.Append(NewAxisId(lineChart, valueAxisId));
            lineChart.Append(NewAxisId(lineChart, seriesAxisId));
            return lineChart;
        }

        private static OfficeXmlNode CreateAreaChart(OfficeXmlNode parent, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors,
            string grouping, uint categoryAxisId, uint valueAxisId) {
            OfficeXmlNode areaChart = parent.New(C, "areaChart");
            areaChart.Append(areaChart.NewVal(C, "grouping", grouping));
            areaChart.Append(areaChart.NewVal(C, "varyColors", false));

            foreach (var descriptor in seriesDescriptors) {
                areaChart.Append(CreateAreaChartSeries(areaChart, descriptor.Index, range, descriptor.Series));
            }

            areaChart.Append(CreateDefaultDataLabels(areaChart));
            areaChart.Append(NewAxisId(areaChart, categoryAxisId));
            areaChart.Append(NewAxisId(areaChart, valueAxisId));
            return areaChart;
        }

        // Area, radar, surface and pie series share one shape: name, categories and values.
        private static OfficeXmlNode CreateCategorySeries(OfficeXmlNode chart, int seriesIndex, ExcelChartDataRange range, ExcelChartSeries? series) {
            OfficeXmlNode element = NewSeries(chart, seriesIndex);
            element.Append(CreateSeriesText(element, range, seriesIndex, series?.Name ?? $"Series {seriesIndex + 1}"));
            element.Append(CreateCategoryAxisData(element, range));
            element.Append(CreateValues(element, range, seriesIndex, series));
            return element;
        }

        private static OfficeXmlNode CreateAreaChartSeries(OfficeXmlNode chart, int seriesIndex, ExcelChartDataRange range, ExcelChartSeries? series) =>
            CreateCategorySeries(chart, seriesIndex, range, series);

        private static OfficeXmlNode CreateArea3DChart(OfficeXmlNode parent, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors,
            string grouping, uint categoryAxisId, uint valueAxisId, uint seriesAxisId) {
            OfficeXmlNode areaChart = parent.New(C, "area3DChart");
            areaChart.Append(areaChart.NewVal(C, "grouping", grouping));
            areaChart.Append(areaChart.NewVal(C, "varyColors", false));

            foreach (var descriptor in seriesDescriptors) {
                areaChart.Append(CreateAreaChartSeries(areaChart, descriptor.Index, range, descriptor.Series));
            }

            areaChart.Append(CreateDefaultDataLabels(areaChart));
            areaChart.Append(areaChart.NewVal(C, "gapDepth", 150U));
            areaChart.Append(NewAxisId(areaChart, categoryAxisId));
            areaChart.Append(NewAxisId(areaChart, valueAxisId));
            areaChart.Append(NewAxisId(areaChart, seriesAxisId));
            return areaChart;
        }

        private static OfficeXmlNode CreateRadarChart(OfficeXmlNode parent, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors,
            uint categoryAxisId, uint valueAxisId) {
            OfficeXmlNode radarChart = parent.New(C, "radarChart");
            radarChart.Append(radarChart.NewVal(C, "radarStyle", "standard"));
            radarChart.Append(radarChart.NewVal(C, "varyColors", false));

            foreach (var descriptor in seriesDescriptors) {
                radarChart.Append(CreateCategorySeries(radarChart, descriptor.Index, range, descriptor.Series));
            }

            radarChart.Append(CreateDefaultDataLabels(radarChart));
            radarChart.Append(NewAxisId(radarChart, categoryAxisId));
            radarChart.Append(NewAxisId(radarChart, valueAxisId));
            return radarChart;
        }

        private static OfficeXmlNode CreateStockChart(OfficeXmlNode parent, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors,
            uint categoryAxisId, uint valueAxisId) {
            if (seriesDescriptors.Count < 3 || seriesDescriptors.Count > 4) {
                throw new ArgumentException("Stock charts require three series (high-low-close) or four series (open-high-low-close).", nameof(seriesDescriptors));
            }

            OfficeXmlNode stockChart = parent.New(C, "stockChart");
            foreach (var descriptor in seriesDescriptors) {
                stockChart.Append(CreateLineChartSeries(stockChart, descriptor.Index, range, descriptor.Series));
            }

            stockChart.Append(CreateDefaultDataLabels(stockChart));
            stockChart.Append(stockChart.New(C, "hiLowLines"));
            if (seriesDescriptors.Count == 4) {
                stockChart.Append(CreateUpDownBars(stockChart, 150U));
            }
            stockChart.Append(NewAxisId(stockChart, categoryAxisId));
            stockChart.Append(NewAxisId(stockChart, valueAxisId));
            return stockChart;
        }

        private static OfficeXmlNode CreateUpDownBars(OfficeXmlNode stockChart, uint gapWidth) {
            OfficeXmlNode bars = stockChart.New(C, "upDownBars");
            bars.Append(bars.NewVal(C, "gapWidth", gapWidth));
            bars.Append(bars.New(C, "upBars"));
            bars.Append(bars.New(C, "downBars"));
            return bars;
        }

        private static OfficeXmlNode CreateSurface3DChart(OfficeXmlNode parent, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors,
            bool wireframe, uint categoryAxisId, uint valueAxisId, uint seriesAxisId) {
            OfficeXmlNode surfaceChart = parent.New(C, "surface3DChart");
            surfaceChart.Append(surfaceChart.NewVal(C, "wireframe", wireframe));
            surfaceChart.Append(surfaceChart.NewVal(C, "varyColors", false));

            foreach (var descriptor in seriesDescriptors) {
                surfaceChart.Append(CreateCategorySeries(surfaceChart, descriptor.Index, range, descriptor.Series));
            }

            surfaceChart.Append(NewAxisId(surfaceChart, categoryAxisId));
            surfaceChart.Append(NewAxisId(surfaceChart, valueAxisId));
            surfaceChart.Append(NewAxisId(surfaceChart, seriesAxisId));
            return surfaceChart;
        }

        private static OfficeXmlNode CreateSurfaceChart(OfficeXmlNode parent, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors,
            bool wireframe, uint categoryAxisId, uint valueAxisId, uint seriesAxisId) {
            OfficeXmlNode surfaceChart = parent.New(C, "surfaceChart");
            surfaceChart.Append(surfaceChart.NewVal(C, "wireframe", wireframe));

            foreach (var descriptor in seriesDescriptors) {
                surfaceChart.Append(CreateCategorySeries(surfaceChart, descriptor.Index, range, descriptor.Series));
            }

            surfaceChart.Append(NewAxisId(surfaceChart, categoryAxisId));
            surfaceChart.Append(NewAxisId(surfaceChart, valueAxisId));
            surfaceChart.Append(NewAxisId(surfaceChart, seriesAxisId));
            return surfaceChart;
        }

        private static OfficeXmlNode CreatePieChart(OfficeXmlNode parent, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors) {
            OfficeXmlNode pieChart = parent.New(C, "pieChart");
            pieChart.Append(pieChart.NewVal(C, "varyColors", true));

            foreach (var descriptor in seriesDescriptors) {
                pieChart.Append(CreateCategorySeries(pieChart, descriptor.Index, range, descriptor.Series));
            }

            pieChart.Append(CreateDefaultDataLabels(pieChart));
            return pieChart;
        }

        private static OfficeXmlNode CreatePie3DChart(OfficeXmlNode parent, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors) {
            OfficeXmlNode pieChart = parent.New(C, "pie3DChart");
            pieChart.Append(pieChart.NewVal(C, "varyColors", true));

            var descriptor = GetSingleSeriesDescriptor(seriesDescriptors, "3-D pie");
            pieChart.Append(CreateCategorySeries(pieChart, descriptor.Index, range, descriptor.Series));

            pieChart.Append(CreateDefaultDataLabels(pieChart));
            return pieChart;
        }

        private static OfficeXmlNode CreateOfPieChart(OfficeXmlNode parent, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors, string type) {
            OfficeXmlNode chart = parent.New(C, "ofPieChart");
            chart.Append(chart.NewVal(C, "ofPieType", type));
            chart.Append(chart.NewVal(C, "varyColors", true));

            string chartName = type == OfPieTypeBar ? "bar-of-pie" : "pie-of-pie";
            var descriptor = GetSingleSeriesDescriptor(seriesDescriptors, chartName);
            chart.Append(CreateCategorySeries(chart, descriptor.Index, range, descriptor.Series));

            chart.Append(CreateDefaultDataLabels(chart));
            chart.Append(chart.NewVal(C, "gapWidth", 150U));
            chart.Append(chart.NewVal(C, "splitType", "pos"));
            chart.Append(chart.NewVal(C, "splitPos", 2D));
            chart.Append(chart.NewVal(C, "secondPieSize", 75U));
            chart.Append(chart.New(C, "serLines"));
            return chart;
        }

        private static SeriesDescriptor GetSingleSeriesDescriptor(IReadOnlyList<SeriesDescriptor> seriesDescriptors, string chartName) {
            if (seriesDescriptors.Count != 1) {
                throw new NotSupportedException($"{chartName} charts support exactly one series.");
            }

            return seriesDescriptors[0];
        }

        private static OfficeXmlNode CreateDoughnutChart(OfficeXmlNode parent, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors) {
            OfficeXmlNode chart = parent.New(C, "doughnutChart");
            chart.Append(chart.NewVal(C, "varyColors", true));

            foreach (var descriptor in seriesDescriptors) {
                chart.Append(CreateCategorySeries(chart, descriptor.Index, range, descriptor.Series));
            }

            chart.Append(CreateDefaultDataLabels(chart));
            chart.Append(chart.NewVal(C, "holeSize", 50));
            return chart;
        }

        private static OfficeXmlNode CreateScatterChart(OfficeXmlNode parent, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> seriesDescriptors,
            uint xAxisId, uint yAxisId, ExcelChartData? data) {
            OfficeXmlNode scatterChart = parent.New(C, "scatterChart");
            scatterChart.Append(scatterChart.NewVal(C, "scatterStyle", "lineMarker"));
            scatterChart.Append(scatterChart.NewVal(C, "varyColors", false));

            IReadOnlyList<double>? xValues = data != null ? ParseNumericCategories(data.Categories) : null;
            foreach (var descriptor in seriesDescriptors) {
                IReadOnlyList<double>? seriesXValues = descriptor.Series?.XValues;
                scatterChart.Append(CreateScatterChartSeries(
                    scatterChart,
                    descriptor.Index,
                    range,
                    descriptor.Series,
                    seriesXValues ?? xValues,
                    ShouldUseLiteralXValues(seriesXValues, xValues)));
            }

            scatterChart.Append(CreateDefaultDataLabels(scatterChart));
            scatterChart.Append(NewAxisId(scatterChart, xAxisId));
            scatterChart.Append(NewAxisId(scatterChart, yAxisId));
            return scatterChart;
        }

        private static OfficeXmlNode CreateScatterChartSeries(OfficeXmlNode chart, int seriesIndex, ExcelChartDataRange range,
            ExcelChartSeries? series, IReadOnlyList<double>? xValues, bool useLiteralXValues) {
            OfficeXmlNode element = NewSeries(chart, seriesIndex);
            element.Append(CreateSeriesText(element, range, seriesIndex, series?.Name ?? $"Series {seriesIndex + 1}"));
            element.Append(CreateXValues(element, range, xValues, useLiteralXValues));
            element.Append(CreateYValues(element, range, seriesIndex, series));
            element.Append(element.NewVal(C, "smooth", false));
            return element;
        }

        private static OfficeXmlNode CreateScatterChartFromRanges(OfficeXmlNode parent, IReadOnlyList<ExcelChartSeriesRange> seriesRanges,
            string defaultSheetName, uint xAxisId, uint yAxisId) {
            OfficeXmlNode scatterChart = parent.New(C, "scatterChart");
            scatterChart.Append(scatterChart.NewVal(C, "scatterStyle", "lineMarker"));
            scatterChart.Append(scatterChart.NewVal(C, "varyColors", false));

            for (int i = 0; i < seriesRanges.Count; i++) {
                scatterChart.Append(CreateScatterChartSeriesFromRanges(scatterChart, i, seriesRanges[i], defaultSheetName));
            }

            scatterChart.Append(CreateDefaultDataLabels(scatterChart));
            scatterChart.Append(NewAxisId(scatterChart, xAxisId));
            scatterChart.Append(NewAxisId(scatterChart, yAxisId));
            return scatterChart;
        }

        private static OfficeXmlNode CreateScatterChartSeriesFromRanges(OfficeXmlNode chart, int seriesIndex, ExcelChartSeriesRange seriesRange,
            string defaultSheetName) {
            string name = string.IsNullOrWhiteSpace(seriesRange.Name) ? $"Series {seriesIndex + 1}" : seriesRange.Name;
            OfficeXmlNode element = NewSeries(chart, seriesIndex);
            element.Append(CreateSeriesTextLiteral(element, name));
            element.Append(CreateNumberRangeSource(element, "xVal", seriesRange.XRangeA1, defaultSheetName));
            element.Append(CreateNumberRangeSource(element, "yVal", seriesRange.YRangeA1, defaultSheetName));
            element.Append(element.NewVal(C, "smooth", false));
            return element;
        }

        private static OfficeXmlNode CreateBubbleChartFromRanges(OfficeXmlNode parent, IReadOnlyList<ExcelChartSeriesRange> seriesRanges,
            string defaultSheetName, uint xAxisId, uint yAxisId) {
            OfficeXmlNode bubbleChart = parent.New(C, "bubbleChart");
            bubbleChart.Append(bubbleChart.NewVal(C, "varyColors", false));

            for (int i = 0; i < seriesRanges.Count; i++) {
                bubbleChart.Append(CreateBubbleChartSeriesFromRanges(bubbleChart, i, seriesRanges[i], defaultSheetName));
            }

            bubbleChart.Append(CreateDefaultDataLabels(bubbleChart));
            bubbleChart.Append(bubbleChart.NewVal(C, "bubbleScale", 100U));
            bubbleChart.Append(bubbleChart.NewVal(C, "showNegBubbles", false));
            bubbleChart.Append(bubbleChart.NewVal(C, "sizeRepresents", "area"));
            bubbleChart.Append(NewAxisId(bubbleChart, xAxisId));
            bubbleChart.Append(NewAxisId(bubbleChart, yAxisId));
            return bubbleChart;
        }

        private static OfficeXmlNode CreateBubbleChartSeriesFromRanges(OfficeXmlNode chart, int seriesIndex, ExcelChartSeriesRange seriesRange,
            string defaultSheetName) {
            if (string.IsNullOrWhiteSpace(seriesRange.BubbleSizeRangeA1)) {
                throw new ArgumentException("Bubble size range is required for bubble charts.", nameof(seriesRange));
            }

            string name = string.IsNullOrWhiteSpace(seriesRange.Name) ? $"Series {seriesIndex + 1}" : seriesRange.Name;
            OfficeXmlNode element = NewSeries(chart, seriesIndex);
            element.Append(CreateSeriesTextLiteral(element, name));
            element.Append(CreateNumberRangeSource(element, "xVal", seriesRange.XRangeA1, defaultSheetName));
            element.Append(CreateNumberRangeSource(element, "yVal", seriesRange.YRangeA1, defaultSheetName));
            element.Append(CreateNumberRangeSource(element, "bubbleSize", seriesRange.BubbleSizeRangeA1!, defaultSheetName));
            return element;
        }

        // A c:xVal, c:yVal or c:bubbleSize holding a number reference to a range, without a cache.
        private static OfficeXmlNode CreateNumberRangeSource(OfficeXmlNode series, string localName, string rangeA1, string defaultSheetName) {
            string formula = EnsureSheetQualifiedRange(defaultSheetName, rangeA1);
            OfficeXmlNode source = series.New(C, localName);
            source.Append(CreateNumberReference(source, formula, null));
            return source;
        }

        private static OfficeXmlNode CreateCategoryAxis(OfficeXmlNode parent, uint axisId, uint crossingAxisId, string? position = null) {
            string axisPosition = position ?? AxisPositionBottom;
            OfficeXmlNode axis = parent.New(C, "catAx");
            AppendAxisStart(axis, axisId, axisPosition);
            AppendNumberFormat(axis);
            AppendAxisTicks(axis, crossingAxisId);
            axis.Append(axis.NewVal(C, "crosses", "autoZero"));
            axis.Append(axis.NewVal(C, "auto", true));
            axis.Append(axis.NewVal(C, "lblAlgn", "ctr"));
            axis.Append(axis.NewVal(C, "lblOffset", 100U));
            axis.Append(axis.NewVal(C, "noMultiLvlLbl", false));
            return axis;
        }

        private static OfficeXmlNode CreateValueAxis(OfficeXmlNode parent, uint axisId, uint crossingAxisId, string? position = null) {
            string axisPosition = position ?? AxisPositionLeft;
            OfficeXmlNode axis = parent.New(C, "valAx");
            AppendAxisStart(axis, axisId, axisPosition);
            axis.Append(axis.New(C, "majorGridlines"));
            AppendNumberFormat(axis);
            AppendAxisTicks(axis, crossingAxisId);
            axis.Append(axis.NewVal(C, "crosses", "autoZero"));
            axis.Append(axis.NewVal(C, "crossBetween", "between"));
            return axis;
        }

        private static OfficeXmlNode CreateSeriesAxis(OfficeXmlNode parent, uint axisId, uint crossingAxisId, string? position = null) {
            string axisPosition = position ?? AxisPositionRight;
            OfficeXmlNode axis = parent.New(C, "serAx");
            AppendAxisStart(axis, axisId, axisPosition);
            AppendAxisTicks(axis, crossingAxisId);
            return axis;
        }

        // c:axId, c:scaling, c:delete and c:axPos, which every axis starts with.
        private static void AppendAxisStart(OfficeXmlNode axis, uint axisId, string axisPosition) {
            axis.Append(NewAxisId(axis, axisId));
            OfficeXmlNode scaling = axis.New(C, "scaling");
            scaling.Append(scaling.NewVal(C, "orientation", "minMax"));
            axis.Append(scaling);
            axis.Append(axis.NewVal(C, "delete", false));
            axis.Append(axis.NewVal(C, "axPos", axisPosition));
        }

        private static void AppendNumberFormat(OfficeXmlNode axis) {
            axis.Append(axis.New(C, "numFmt").Attr("formatCode", "General").Attr("sourceLinked", true));
        }

        // The tick marks, tick label position and crossing axis id.
        private static void AppendAxisTicks(OfficeXmlNode axis, uint crossingAxisId) {
            axis.Append(axis.NewVal(C, "majorTickMark", "none"));
            axis.Append(axis.NewVal(C, "minorTickMark", "none"));
            axis.Append(axis.NewVal(C, "tickLblPos", "nextTo"));
            axis.Append(axis.NewVal(C, "crossAx", crossingAxisId));
        }

        private static OfficeXmlNode CreateDefaultDataLabels(OfficeXmlNode parent) {
            OfficeXmlNode labels = parent.New(C, "dLbls");
            labels.Append(labels.NewVal(C, "showLegendKey", false));
            labels.Append(labels.NewVal(C, "showVal", false));
            labels.Append(labels.NewVal(C, "showCatName", false));
            labels.Append(labels.NewVal(C, "showSerName", false));
            labels.Append(labels.NewVal(C, "showPercent", false));
            labels.Append(labels.NewVal(C, "showBubbleSize", false));
            return labels;
        }

        private static OfficeXmlNode CreateSeriesText(OfficeXmlNode series, ExcelChartDataRange range, int seriesIndex, string name) {
            if (range.HasHeaderRow) {
                string seriesCell = range.SeriesNameCellA1(seriesIndex);
                string formula = BuildSheetQualifiedRange(range.SheetName, seriesCell);
                OfficeXmlNode text = series.New(C, "tx");
                text.Append(CreateSingleStringReference(text, formula, name));
                return text;
            }

            return CreateSeriesTextLiteral(series, name);
        }

        private static OfficeXmlNode CreateSeriesTextLiteral(OfficeXmlNode series, string name) {
            OfficeXmlNode text = series.New(C, "tx");
            text.Append(text.NewText(C, "v", name));
            return text;
        }

        private static OfficeXmlNode CreateCategoryAxisData(OfficeXmlNode series, ExcelChartDataRange range) {
            string formula = BuildSheetQualifiedRange(range.SheetName, range.CategoriesRangeA1);
            OfficeXmlNode categories = series.New(C, "cat");
            categories.Append(CreateStringReference(categories, formula, null));
            return categories;
        }

        private static bool ShouldUseLiteralXValues(IReadOnlyList<double>? seriesXValues, IReadOnlyList<double>? categoryXValues) {
            if (seriesXValues == null || categoryXValues == null || seriesXValues.Count != categoryXValues.Count) {
                return seriesXValues != null;
            }

            for (int i = 0; i < seriesXValues.Count; i++) {
                if (!seriesXValues[i].Equals(categoryXValues[i])) {
                    return true;
                }
            }

            return false;
        }

        private static OfficeXmlNode CreateXValues(OfficeXmlNode series, ExcelChartDataRange range, IReadOnlyList<double>? xValues, bool useLiteralXValues = false) {
            OfficeXmlNode element = series.New(C, "xVal");
            if (useLiteralXValues && xValues != null) {
                element.Append(CreateNumberLiteral(element, xValues));
                return element;
            }

            string formula = BuildSheetQualifiedRange(range.SheetName, range.CategoriesRangeA1);
            element.Append(CreateNumberReference(element, formula, xValues));
            return element;
        }

        private static OfficeXmlNode CreateValues(OfficeXmlNode seriesElement, ExcelChartDataRange range, int seriesIndex, ExcelChartSeries? series) {
            string formula = BuildSheetQualifiedRange(range.SheetName, range.SeriesValuesRangeA1(seriesIndex));
            OfficeXmlNode element = seriesElement.New(C, "val");
            element.Append(CreateNumberReference(element, formula, series?.Values));
            return element;
        }

        private static OfficeXmlNode CreateYValues(OfficeXmlNode seriesElement, ExcelChartDataRange range, int seriesIndex, ExcelChartSeries? series) {
            string formula = BuildSheetQualifiedRange(range.SheetName, range.SeriesValuesRangeA1(seriesIndex));
            OfficeXmlNode element = seriesElement.New(C, "yVal");
            element.Append(CreateNumberReference(element, formula, series?.Values));
            return element;
        }
    }
}
