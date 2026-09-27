using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartWriter {
        private static C.BarChartSeries CreateBarChartSeries(int seriesIndex, OfficeChartSeries series, IReadOnlyList<string> categories) {
            int lastRow = categories.Count + 1;
            string seriesColumn = ColumnLetter(seriesIndex + 2);
            string seriesNameRef = $"Sheet1!${seriesColumn}$1";
            string categoriesRef = $"Sheet1!$A$2:$A${lastRow}";
            string valuesRef = $"Sheet1!${seriesColumn}$2:${seriesColumn}${lastRow}";

            C.BarChartSeries seriesElement = new(
                new C.Index { Val = (uint)seriesIndex },
                new C.Order { Val = (uint)seriesIndex },
                new C.SeriesText(CreateStringReference(seriesNameRef, new[] { series.Name })),
                new C.InvertIfNegative { Val = false },
                new C.CategoryAxisData(CreateStringReference(categoriesRef, categories)),
                new C.Values(CreateNumberReference(valuesRef, series.Values))
            );

            return seriesElement;
        }

        private static C.LineChartSeries CreateLineChartSeries(int seriesIndex, OfficeChartSeries series, IReadOnlyList<string> categories) {
            int lastRow = categories.Count + 1;
            string seriesColumn = ColumnLetter(seriesIndex + 2);
            string seriesNameRef = $"Sheet1!${seriesColumn}$1";
            string categoriesRef = $"Sheet1!$A$2:$A${lastRow}";
            string valuesRef = $"Sheet1!${seriesColumn}$2:${seriesColumn}${lastRow}";

            C.LineChartSeries seriesElement = new(
                new C.Index { Val = (uint)seriesIndex },
                new C.Order { Val = (uint)seriesIndex },
                new C.SeriesText(CreateStringReference(seriesNameRef, new[] { series.Name })),
                new C.CategoryAxisData(CreateStringReference(categoriesRef, categories)),
                new C.Values(CreateNumberReference(valuesRef, series.Values))
            );

            return seriesElement;
        }

        private static C.ScatterChart CreateScatterChart(OfficeChartData data, out uint xAxisId, out uint yAxisId) {
            xAxisId = GetNextAxisId();
            yAxisId = GetNextAxisId();

            C.ScatterChart scatterChart = new(
                new C.ScatterStyle { Val = C.ScatterStyleValues.LineMarker },
                new C.VaryColors { Val = false });

            for (int i = 0; i < data.Series.Count; i++) {
                scatterChart.Append(CreateScatterChartSeries(i, data.Series[i]));
            }

            scatterChart.Append(CreateDefaultDataLabels());
            scatterChart.Append(new C.AxisId { Val = xAxisId });
            scatterChart.Append(new C.AxisId { Val = yAxisId });
            return scatterChart;
        }

        private static C.ScatterChartSeries CreateScatterChartSeries(int seriesIndex, OfficeChartSeries series) {
            string seriesNameRef = GetScatterSeriesNameReference(seriesIndex);
            string xValuesRef = GetScatterXValuesReference(seriesIndex, series.XValues!.Count);
            string yValuesRef = GetScatterYValuesReference(seriesIndex, series.Values.Count);

            C.ScatterChartSeries seriesElement = new(
                new C.Index { Val = (uint)seriesIndex },
                new C.Order { Val = (uint)seriesIndex },
                new C.SeriesText(CreateStringReference(seriesNameRef, new[] { series.Name })),
                new C.XValues(CreateNumberReference(xValuesRef, series.XValues!)),
                new C.YValues(CreateNumberReference(yValuesRef, series.Values))
            );

            return seriesElement;
        }

        private static C.PieChart CreatePieChart(OfficeChartData data) {
            C.PieChart pieChart = new(
                new C.VaryColors { Val = true });

            for (int i = 0; i < data.Series.Count; i++) {
                pieChart.Append(CreatePieChartSeries(i, data.Series[i], data.Categories));
            }

            pieChart.Append(CreateDefaultDataLabels());
            pieChart.Append(new C.FirstSliceAngle { Val = (UInt16Value)0U });
            return pieChart;
        }

        private static C.DoughnutChart CreateDoughnutChart(OfficeChartData data) {
            C.DoughnutChart doughnutChart = new(
                new C.VaryColors { Val = true });

            for (int i = 0; i < data.Series.Count; i++) {
                doughnutChart.Append(CreatePieChartSeries(i, data.Series[i], data.Categories));
            }

            doughnutChart.Append(CreateDefaultDataLabels());
            doughnutChart.Append(new C.FirstSliceAngle { Val = (UInt16Value)0U });
            doughnutChart.Append(new C.HoleSize { Val = (ByteValue)50 });
            return doughnutChart;
        }

        private static C.PieChartSeries CreatePieChartSeries(int seriesIndex, OfficeChartSeries series, IReadOnlyList<string> categories) {
            int lastRow = categories.Count + 1;
            string seriesColumn = ColumnLetter(seriesIndex + 2);
            string seriesNameRef = $"Sheet1!${seriesColumn}$1";
            string categoriesRef = $"Sheet1!$A$2:$A${lastRow}";
            string valuesRef = $"Sheet1!${seriesColumn}$2:${seriesColumn}${lastRow}";

            C.PieChartSeries seriesElement = new(
                new C.Index { Val = (uint)seriesIndex },
                new C.Order { Val = (uint)seriesIndex },
                new C.SeriesText(CreateStringReference(seriesNameRef, new[] { series.Name })),
                new C.CategoryAxisData(CreateStringReference(categoriesRef, categories)),
                new C.Values(CreateNumberReference(valuesRef, series.Values))
            );

            return seriesElement;
        }

        private static C.ValueAxis CreateValueAxis(uint axisId, uint crossingAxisId) {
            return CreateValueAxis(axisId, crossingAxisId, C.AxisPositionValues.Left);
        }

        private static C.ValueAxis CreateValueAxis(uint axisId, uint crossingAxisId,
            C.AxisPositionValues axisPosition) {
            C.ValueAxis axis = new(
                new C.AxisId { Val = axisId },
                new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new C.Delete { Val = false },
                new C.AxisPosition { Val = axisPosition },
                new C.MajorGridlines(),
                new C.NumberingFormat { FormatCode = "General", SourceLinked = true },
                new C.MajorTickMark { Val = C.TickMarkValues.None },
                new C.MinorTickMark { Val = C.TickMarkValues.None },
                new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
                new C.CrossingAxis { Val = crossingAxisId },
                new C.Crosses { Val = C.CrossesValues.AutoZero },
                new C.CrossBetween { Val = C.CrossBetweenValues.Between }
            );

            return axis;
        }

        private static C.DataLabels CreateDefaultDataLabels() {
            return new C.DataLabels(
                new C.ShowLegendKey { Val = false },
                new C.ShowValue { Val = false },
                new C.ShowCategoryName { Val = false },
                new C.ShowSeriesName { Val = false },
                new C.ShowPercent { Val = false },
                new C.ShowBubbleSize { Val = false }
            );
        }

        private static C.StringReference CreateStringReference(string formula, IReadOnlyList<string> values) {
            C.StringCache cache = new();
            cache.Append(new C.PointCount { Val = (uint)values.Count });
            for (int i = 0; i < values.Count; i++) {
                cache.Append(new C.StringPoint {
                    Index = (uint)i,
                    NumericValue = new C.NumericValue { Text = values[i] ?? string.Empty }
                });
            }

            return new C.StringReference(
                new C.Formula { Text = formula },
                cache);
        }

        private static C.NumberReference CreateNumberReference(string formula, IReadOnlyList<double> values) {
            C.NumberingCache cache = new();
            cache.Append(new C.FormatCode { Text = "General" });
            cache.Append(new C.PointCount { Val = (uint)values.Count });
            for (int i = 0; i < values.Count; i++) {
                cache.Append(new C.NumericPoint {
                    Index = (uint)i,
                    NumericValue = new C.NumericValue { Text = values[i].ToString(CultureInfo.InvariantCulture) }
                });
            }

            return new C.NumberReference(
                new C.Formula { Text = formula },
                cache);
        }

        private static void UpdateSeriesIndexOrder(OpenXmlCompositeElement series, int index) {
            C.Index idx = series.GetFirstChild<C.Index>() ?? new C.Index();
            idx.Val = (uint)index;
            if (idx.Parent == null) {
                series.PrependChild(idx);
            }

            C.Order order = series.GetFirstChild<C.Order>() ?? new C.Order();
            order.Val = (uint)index;
            if (order.Parent == null) {
                series.InsertAfter(order, idx);
            }
        }

        private static void UpdateScatterSeriesText(C.ScatterChartSeries series, int seriesIndex, string seriesName) {
            string seriesNameRef = GetScatterSeriesNameReference(seriesIndex);
            C.SeriesText seriesText = series.GetFirstChild<C.SeriesText>() ?? new C.SeriesText();
            seriesText.RemoveAllChildren<C.StringReference>();
            seriesText.RemoveAllChildren<C.StringLiteral>();
            seriesText.Append(CreateStringReference(seriesNameRef, new[] { seriesName }));

            if (seriesText.Parent == null) {
                OpenXmlElement? insertAfter = series.GetFirstChild<C.Order>();
                insertAfter ??= series.GetFirstChild<C.Index>();
                if (insertAfter != null) {
                    series.InsertAfter(seriesText, insertAfter);
                } else {
                    series.PrependChild(seriesText);
                }
            }
        }

        private static void UpdateXValues(C.ScatterChartSeries series, int seriesIndex, IReadOnlyList<double> values) {
            string valuesRef = GetScatterXValuesReference(seriesIndex, values.Count);
            C.XValues xValueElement = series.GetFirstChild<C.XValues>() ?? new C.XValues();
            xValueElement.RemoveAllChildren<C.NumberReference>();
            xValueElement.RemoveAllChildren<C.NumberLiteral>();
            xValueElement.Append(CreateNumberReference(valuesRef, values));

            if (xValueElement.Parent == null) {
                series.Append(xValueElement);
            }
        }

        private static void UpdateYValues(C.ScatterChartSeries series, int seriesIndex, IReadOnlyList<double> values) {
            string valuesRef = GetScatterYValuesReference(seriesIndex, values.Count);
            C.YValues yValueElement = series.GetFirstChild<C.YValues>() ?? new C.YValues();
            yValueElement.RemoveAllChildren<C.NumberReference>();
            yValueElement.RemoveAllChildren<C.NumberLiteral>();
            yValueElement.Append(CreateNumberReference(valuesRef, values));

            if (yValueElement.Parent == null) {
                series.Append(yValueElement);
            }
        }

        private static void InsertSeries(OpenXmlCompositeElement chart, OpenXmlElement series) {
            OpenXmlElement? insertBefore = chart.ChildElements.FirstOrDefault(child =>
                child is C.DataLabels ||
                child is C.GapWidth ||
                child is C.Overlap ||
                child is C.AxisId ||
                child is C.FirstSliceAngle ||
                child is C.HoleSize ||
                child is C.Marker ||
                child is C.Smooth ||
                child is C.SeriesLines);

            if (insertBefore != null) {
                chart.InsertBefore(series, insertBefore);
            } else {
                chart.Append(series);
            }
        }

        private static string GetScatterSeriesNameReference(int seriesIndex) {
            string yColumn = ColumnLetter((seriesIndex * 2) + 2);
            return $"Sheet1!${yColumn}$1";
        }

        private static string GetScatterXValuesReference(int seriesIndex, int pointCount) {
            string xColumn = ColumnLetter((seriesIndex * 2) + 1);
            int lastRow = pointCount + 1;
            return $"Sheet1!${xColumn}$2:${xColumn}${lastRow}";
        }

        private static string GetScatterYValuesReference(int seriesIndex, int pointCount) {
            string yColumn = ColumnLetter((seriesIndex * 2) + 2);
            int lastRow = pointCount + 1;
            return $"Sheet1!${yColumn}$2:${yColumn}${lastRow}";
        }

        internal static void UpdateScatterData(ChartPart chartPart, OfficeChartData data) {
            if (chartPart == null) {
                throw new ArgumentNullException(nameof(chartPart));
            }
            if (data == null) {
                throw new ArgumentNullException(nameof(data));
            }

            C.ChartSpace? chartSpace = chartPart.ChartSpace;
            C.Chart? chart = chartSpace?.GetFirstChild<C.Chart>();
            C.PlotArea? plotArea = chart?.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                throw new InvalidOperationException("Chart plot area not found.");
            }

            List<C.ScatterChart> scatterCharts = plotArea.Elements<C.ScatterChart>().ToList();
            if (scatterCharts.Count > 0) {
                UpdateScatterChartLayers(plotArea, scatterCharts, data);
                return;
            }

            throw new NotSupportedException("Chart type is not supported for scatter data updates.");
        }
    }
}
