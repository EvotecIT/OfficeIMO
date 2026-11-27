using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    internal static partial class ExcelChartUtils {
        private static void UpdateSeriesIndexOrder(OfficeXmlNode series, int index) {
            OfficeXmlNode idx = series.Child(C, "idx");
            if (idx.IsNone) {
                idx = series.InsertBefore(series.New(C, "idx"), series.FirstChild);
            }
            idx.SetVal((uint)index);

            OfficeXmlNode order = series.Child(C, "order");
            if (order.IsNone) {
                order = idx.InsertAfterSelf(series.New(C, "order"));
            }
            order.SetVal((uint)index);
        }

        private static int GetSeriesIndex(OfficeXmlNode series) {
            return (int)(series.Child(C, "idx").UInt32Val() ?? 0U);
        }

        private static HashSet<int> CreateDescriptorIndexSet(IReadOnlyList<SeriesDescriptor> descriptors) {
            var indexes = new HashSet<int>();
            for (int i = 0; i < descriptors.Count; i++) {
                indexes.Add(descriptors[i].Index);
            }

            return indexes;
        }

        private static Dictionary<int, OfficeXmlNode> CreateSeriesIndexMap(IReadOnlyList<OfficeXmlNode> series) {
            var map = new Dictionary<int, OfficeXmlNode>(series.Count);
            for (int i = 0; i < series.Count; i++) {
                map.Add(GetSeriesIndex(series[i]), series[i]);
            }

            return map;
        }

        // The child order of a c:ser, which depends on the chart element holding it.
        private static OfficeXmlChildOrder SeriesOrder(OfficeXmlNode series) {
            switch (series.Parent.LocalName) {
                case "barChart":
                case "bar3DChart":
                    return CBarSerView.Order;
                case "lineChart":
                case "line3DChart":
                case "stockChart":
                    return CLineSerView.Order;
                case "areaChart":
                case "area3DChart":
                    return CAreaSerView.Order;
                case "pieChart":
                case "pie3DChart":
                case "ofPieChart":
                case "doughnutChart":
                    return CPieSerView.Order;
                case "radarChart":
                    return CRadarSerView.Order;
                case "surfaceChart":
                case "surface3DChart":
                    return CSurfaceSerView.Order;
                case "scatterChart":
                    return CScatterSerView.Order;
                default:
                    return CBubbleSerView.Order;
            }
        }

        // Puts a series child in schema order, replacing the one already there, as the SDK's AddChild does for an
        // element the series holds once.
        private static void SetSeriesChild(OfficeXmlNode series, OfficeXmlNode child) => series.SetOrdered(child, SeriesOrder(series));

        private static void UpdateSeriesText(OfficeXmlNode series, ExcelChartDataRange range, int seriesIndex, string seriesName) {
            OfficeXmlNode text = series.New(C, "tx");
            if (range.HasHeaderRow) {
                text.Append(CreateSingleStringReference(text, BuildSheetQualifiedRange(range.SheetName, range.SeriesNameCellA1(seriesIndex)), seriesName));
            } else {
                text.Append(text.NewText(C, "v", seriesName));
            }
            SetSeriesChild(series, text);
        }

        private static void UpdateCategoryAxisData(OfficeXmlNode series, ExcelChartDataRange range, IReadOnlyList<string> categories) {
            string formula = BuildSheetQualifiedRange(range.SheetName, range.CategoriesRangeA1);
            OfficeXmlNode element = series.New(C, "cat");
            element.Append(CreateStringReference(element, formula, categories));
            SetSeriesChild(series, element);
        }

        private static void UpdateValues(OfficeXmlNode series, ExcelChartDataRange range, int seriesIndex, IReadOnlyList<double> values) {
            string formula = BuildSheetQualifiedRange(range.SheetName, range.SeriesValuesRangeA1(seriesIndex));
            OfficeXmlNode element = series.New(C, "val");
            element.Append(CreateNumberReference(element, formula, values));
            SetSeriesChild(series, element);
        }

        private static void UpdateXValues(OfficeXmlNode series, ExcelChartDataRange range, IReadOnlyList<double> xValues, bool useLiteralXValues = false) {
            string formula = BuildSheetQualifiedRange(range.SheetName, range.CategoriesRangeA1);
            OfficeXmlNode element = series.New(C, "xVal");
            element.Append(useLiteralXValues ? CreateNumberLiteral(element, xValues) : CreateNumberReference(element, formula, xValues));
            SetSeriesChild(series, element);
        }

        private static void UpdateYValues(OfficeXmlNode series, ExcelChartDataRange range, int seriesIndex, IReadOnlyList<double> values) {
            string formula = BuildSheetQualifiedRange(range.SheetName, range.SeriesValuesRangeA1(seriesIndex));
            OfficeXmlNode element = series.New(C, "yVal");
            element.Append(CreateNumberReference(element, formula, values));
            SetSeriesChild(series, element);
        }

        // The chart element children a new series goes before. A chart element's own c:marker (a line chart's marker
        // flag) is not one of them: only a series marker was, and no chart element holds one.
        private static readonly string[] SeriesInsertBeforeNames = {
            "dLbls", "gapWidth", "gapDepth", "overlap", "hiLowLines", "upDownBars", "bandFmts", "shape", "axId", "serLines"
        };

        private static OfficeXmlNode InsertSeries(OfficeXmlNode chart, OfficeXmlNode series) {
            OfficeXmlNode insertBefore = OfficeXmlNode.None;
            foreach (OfficeXmlNode child in chart.Children()) {
                if (child.NamespaceId == C && Array.IndexOf(SeriesInsertBeforeNames, child.LocalName) >= 0) {
                    insertBefore = child;
                    break;
                }
            }

            return chart.InsertBefore(series, insertBefore);
        }

        private static void EnsureStockChartLines(OfficeXmlNode stockChart, int seriesCount) {
            if (stockChart.Child(C, "hiLowLines").IsNone) {
                OfficeXmlNode insertBefore = stockChart.Child(C, "upDownBars");
                if (insertBefore.IsNone) insertBefore = stockChart.Child(C, "axId");
                stockChart.InsertBefore(stockChart.New(C, "hiLowLines"), insertBefore);
            }

            OfficeXmlNode bars = stockChart.Child(C, "upDownBars");
            if (seriesCount == 4) {
                if (bars.IsNone) {
                    bars = CreateUpDownBars(stockChart, (uint)OfficeTheme.Default.ChartStyle.BarGapPercent);
                    stockChart.InsertBefore(bars, stockChart.Child(C, "axId"));
                }
            } else if (bars.Exists) {
                bars.Remove();
            }
        }
    }
}
