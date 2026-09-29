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

        private static HashSet<int> CreateDescriptorIndexSet(IReadOnlyList<SeriesDescriptor> descriptors) {
            var indexes = new HashSet<int>();
            for (int i = 0; i < descriptors.Count; i++) {
                indexes.Add(descriptors[i].Index);
            }

            return indexes;
        }

        private static Dictionary<int, TSeries> CreateSeriesIndexMap<TSeries>(IReadOnlyList<TSeries> series)
            where TSeries : OpenXmlCompositeElement {
            var map = new Dictionary<int, TSeries>(series.Count);
            for (int i = 0; i < series.Count; i++) {
                map.Add(GetSeriesIndex(series[i]), series[i]);
            }

            return map;
        }

        private static void UpdateSeriesText(OpenXmlCompositeElement series, ExcelChartDataRange range, int seriesIndex, string seriesName) {
            SeriesText text = range.HasHeaderRow
                ? new SeriesText(CreateSingleStringReference(BuildSheetQualifiedRange(range.SheetName, range.SeriesNameCellA1(seriesIndex)), seriesName))
                : new SeriesText(new NumericValue { Text = seriesName });
            series.AddChild(text, true);
        }

        private static void UpdateCategoryAxisData(OpenXmlCompositeElement series, ExcelChartDataRange range, IReadOnlyList<string> categories) {
            string formula = BuildSheetQualifiedRange(range.SheetName, range.CategoriesRangeA1);
            series.AddChild(new CategoryAxisData(CreateStringReference(formula, categories)), true);
        }

        private static void UpdateValues(OpenXmlCompositeElement series, ExcelChartDataRange range, int seriesIndex, IReadOnlyList<double> values) {
            string formula = BuildSheetQualifiedRange(range.SheetName, range.SeriesValuesRangeA1(seriesIndex));
            series.AddChild(new Values(CreateNumberReference(formula, values)), true);
        }

        private static void UpdateXValues(ScatterChartSeries series, ExcelChartDataRange range, IReadOnlyList<double> xValues, bool useLiteralXValues = false) {
            string formula = BuildSheetQualifiedRange(range.SheetName, range.CategoriesRangeA1);
            series.AddChild(new XValues(useLiteralXValues ? CreateNumberLiteral(xValues) : CreateNumberReference(formula, xValues)), true);
        }

        private static void UpdateYValues(ScatterChartSeries series, ExcelChartDataRange range, int seriesIndex, IReadOnlyList<double> values) {
            string formula = BuildSheetQualifiedRange(range.SheetName, range.SeriesValuesRangeA1(seriesIndex));
            series.AddChild(new YValues(CreateNumberReference(formula, values)), true);
        }
        private static void InsertSeries(OpenXmlCompositeElement chart, OpenXmlElement series) {
            OpenXmlElement? insertBefore = chart.ChildElements.FirstOrDefault(child =>
                child is DataLabels ||
                child is GapWidth ||
                child is GapDepth ||
                child is Overlap ||
                child is HighLowLines ||
                child is UpDownBars ||
                child is BandFormats ||
                child is Shape ||
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

        private static void EnsureStockChartLines(StockChart stockChart, int seriesCount) {
            if (stockChart.GetFirstChild<HighLowLines>() == null) {
                OpenXmlElement? insertBefore = stockChart.GetFirstChild<UpDownBars>();
                insertBefore ??= stockChart.GetFirstChild<AxisId>();
                if (insertBefore != null) {
                    stockChart.InsertBefore(new HighLowLines(), insertBefore);
                } else {
                    stockChart.Append(new HighLowLines());
                }
            }

            UpDownBars? bars = stockChart.GetFirstChild<UpDownBars>();
            if (seriesCount == 4) {
                if (bars == null) {
                    bars = new UpDownBars(
                        new GapWidth { Val = (UInt16Value)150U },
                        new UpBars(),
                        new DownBars());
                    OpenXmlElement? insertBefore = stockChart.GetFirstChild<AxisId>();
                    if (insertBefore != null) {
                        stockChart.InsertBefore(bars, insertBefore);
                    } else {
                        stockChart.Append(bars);
                    }
                }
            } else {
                bars?.Remove();
            }
        }
    }
}
