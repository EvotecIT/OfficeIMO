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
            SeriesText seriesText = series.GetFirstChild<SeriesText>() ?? new SeriesText();
            seriesText.RemoveAllChildren<StringReference>();
            seriesText.RemoveAllChildren<StringLiteral>();
            seriesText.RemoveAllChildren<NumericValue>();

            if (range.HasHeaderRow) {
                string seriesCell = range.SeriesNameCellA1(seriesIndex);
                string formula = BuildSheetQualifiedRange(range.SheetName, seriesCell);
                seriesText.Append(CreateSingleStringReference(formula, seriesName));
            } else {
                seriesText.Append(new NumericValue { Text = seriesName });
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

        private static void UpdateXValues(ScatterChartSeries series, ExcelChartDataRange range, IReadOnlyList<double> xValues, bool useLiteralXValues = false) {
            string formula = BuildSheetQualifiedRange(range.SheetName, range.CategoriesRangeA1);
            XValues xValueElement = series.GetFirstChild<XValues>() ?? new XValues();
            xValueElement.RemoveAllChildren<NumberReference>();
            xValueElement.RemoveAllChildren<NumberLiteral>();
            xValueElement.Append(useLiteralXValues ? CreateNumberLiteral(xValues) : CreateNumberReference(formula, xValues));

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
