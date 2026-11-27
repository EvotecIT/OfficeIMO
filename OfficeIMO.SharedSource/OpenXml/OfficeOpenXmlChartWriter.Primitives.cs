using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartWriter {
        private static OfficeXmlNode CreateBarChartSeries(OfficeXmlNode context, int seriesIndex, OfficeSeries series, IReadOnlyList<string> categories) {
            int lastRow = categories.Count + 1;
            string seriesColumn = ColumnLetter(seriesIndex + 2);
            string seriesNameRef = $"Sheet1!${seriesColumn}$1";
            string categoriesRef = $"Sheet1!$A$2:$A${lastRow}";
            string valuesRef = $"Sheet1!${seriesColumn}$2:${seriesColumn}${lastRow}";

            OfficeXmlNode seriesElement = context.CreateElement(NsC, "ser");
            seriesElement.AppendVal(NsC, "idx", Text((uint)seriesIndex));
            seriesElement.AppendVal(NsC, "order", Text((uint)seriesIndex));
            AppendStringReference(seriesElement.AppendNew(NsC, "tx"), seriesNameRef, new[] { series.Name });
            seriesElement.AppendVal(NsC, "invertIfNegative", Text(false));
            AppendStringReference(seriesElement.AppendNew(NsC, "cat"), categoriesRef, categories);
            AppendNumberReference(seriesElement.AppendNew(NsC, "val"), valuesRef, series.Values);
            return seriesElement;
        }

        private static OfficeXmlNode CreateLineChartSeries(OfficeXmlNode context, int seriesIndex, OfficeSeries series, IReadOnlyList<string> categories) {
            int lastRow = categories.Count + 1;
            string seriesColumn = ColumnLetter(seriesIndex + 2);
            string seriesNameRef = $"Sheet1!${seriesColumn}$1";
            string categoriesRef = $"Sheet1!$A$2:$A${lastRow}";
            string valuesRef = $"Sheet1!${seriesColumn}$2:${seriesColumn}${lastRow}";

            OfficeXmlNode seriesElement = context.CreateElement(NsC, "ser");
            seriesElement.AppendVal(NsC, "idx", Text((uint)seriesIndex));
            seriesElement.AppendVal(NsC, "order", Text((uint)seriesIndex));
            AppendStringReference(seriesElement.AppendNew(NsC, "tx"), seriesNameRef, new[] { series.Name });
            AppendStringReference(seriesElement.AppendNew(NsC, "cat"), categoriesRef, categories);
            AppendNumberReference(seriesElement.AppendNew(NsC, "val"), valuesRef, series.Values);
            return seriesElement;
        }

        private static OfficeXmlNode CreateScatterChart(OfficeXmlNode context, OfficeChartData data, out uint xAxisId, out uint yAxisId) {
            xAxisId = GetNextAxisId();
            yAxisId = GetNextAxisId();

            OfficeXmlNode scatterChart = context.CreateElement(NsC, "scatterChart");
            scatterChart.AppendVal(NsC, "scatterStyle", "lineMarker");
            scatterChart.AppendVal(NsC, "varyColors", Text(false));

            for (int i = 0; i < data.Series.Count; i++) {
                scatterChart.Append(CreateScatterChartSeries(scatterChart, i, data.Series[i]));
            }

            scatterChart.Append(CreateDefaultDataLabels(scatterChart));
            scatterChart.AppendVal(NsC, "axId", Text(xAxisId));
            scatterChart.AppendVal(NsC, "axId", Text(yAxisId));
            return scatterChart;
        }

        private static OfficeXmlNode CreateScatterChartSeries(OfficeXmlNode context, int seriesIndex, OfficeSeries series) {
            string seriesNameRef = GetScatterSeriesNameReference(seriesIndex);
            string xValuesRef = GetScatterXValuesReference(seriesIndex, series.XValues!.Count);
            string yValuesRef = GetScatterYValuesReference(seriesIndex, series.Values.Count);

            OfficeXmlNode seriesElement = context.CreateElement(NsC, "ser");
            seriesElement.AppendVal(NsC, "idx", Text((uint)seriesIndex));
            seriesElement.AppendVal(NsC, "order", Text((uint)seriesIndex));
            AppendStringReference(seriesElement.AppendNew(NsC, "tx"), seriesNameRef, new[] { series.Name });
            AppendNumberReference(seriesElement.AppendNew(NsC, "xVal"), xValuesRef, series.XValues!);
            AppendNumberReference(seriesElement.AppendNew(NsC, "yVal"), yValuesRef, series.Values);
            return seriesElement;
        }

        private static OfficeXmlNode CreatePieChart(OfficeXmlNode context, OfficeChartData data) {
            OfficeXmlNode pieChart = context.CreateElement(NsC, "pieChart");
            pieChart.AppendVal(NsC, "varyColors", Text(true));

            for (int i = 0; i < data.Series.Count; i++) {
                pieChart.Append(CreatePieChartSeries(pieChart, i, data.Series[i], data.Categories));
            }

            pieChart.Append(CreateDefaultDataLabels(pieChart));
            pieChart.AppendVal(NsC, "firstSliceAng", "0");
            return pieChart;
        }

        private static OfficeXmlNode CreateDoughnutChart(OfficeXmlNode context, OfficeChartData data) {
            OfficeXmlNode doughnutChart = context.CreateElement(NsC, "doughnutChart");
            doughnutChart.AppendVal(NsC, "varyColors", Text(true));

            for (int i = 0; i < data.Series.Count; i++) {
                doughnutChart.Append(CreatePieChartSeries(doughnutChart, i, data.Series[i], data.Categories));
            }

            doughnutChart.Append(CreateDefaultDataLabels(doughnutChart));
            doughnutChart.AppendVal(NsC, "firstSliceAng", "0");
            doughnutChart.AppendVal(NsC, "holeSize", "50");
            return doughnutChart;
        }

        private static OfficeXmlNode CreatePieChartSeries(OfficeXmlNode context, int seriesIndex, OfficeSeries series, IReadOnlyList<string> categories) {
            int lastRow = categories.Count + 1;
            string seriesColumn = ColumnLetter(seriesIndex + 2);
            string seriesNameRef = $"Sheet1!${seriesColumn}$1";
            string categoriesRef = $"Sheet1!$A$2:$A${lastRow}";
            string valuesRef = $"Sheet1!${seriesColumn}$2:${seriesColumn}${lastRow}";

            OfficeXmlNode seriesElement = context.CreateElement(NsC, "ser");
            seriesElement.AppendVal(NsC, "idx", Text((uint)seriesIndex));
            seriesElement.AppendVal(NsC, "order", Text((uint)seriesIndex));
            AppendStringReference(seriesElement.AppendNew(NsC, "tx"), seriesNameRef, new[] { series.Name });
            AppendStringReference(seriesElement.AppendNew(NsC, "cat"), categoriesRef, categories);
            AppendNumberReference(seriesElement.AppendNew(NsC, "val"), valuesRef, series.Values);
            return seriesElement;
        }

        private static OfficeXmlNode CreateValueAxis(OfficeXmlNode context, uint axisId, uint crossingAxisId) {
            return CreateValueAxis(context, axisId, crossingAxisId, "l");
        }

        /// <param name="context">The element the axis is created for, normally the plot area it goes into.</param>
        /// <param name="axisId">The axis id.</param>
        /// <param name="crossingAxisId">The id of the axis this one crosses.</param>
        /// <param name="axisPosition">The <c>c:axPos</c> token: "l", "r", "t" or "b".</param>
        private static OfficeXmlNode CreateValueAxis(OfficeXmlNode context, uint axisId, uint crossingAxisId, string axisPosition) {
            OfficeXmlNode axis = context.CreateElement(NsC, "valAx");
            axis.AppendVal(NsC, "axId", Text(axisId));
            axis.AppendNew(NsC, "scaling").AppendVal(NsC, "orientation", "minMax");
            axis.AppendVal(NsC, "delete", Text(false));
            axis.AppendVal(NsC, "axPos", axisPosition);
            axis.AppendNew(NsC, "majorGridlines");
            OfficeXmlNode format = axis.AppendNew(NsC, "numFmt");
            format.SetAttribute("formatCode", "General");
            format.SetAttribute("sourceLinked", Text(false));
            axis.AppendVal(NsC, "majorTickMark", "none");
            axis.AppendVal(NsC, "minorTickMark", "none");
            axis.AppendVal(NsC, "tickLblPos", "nextTo");
            axis.AppendVal(NsC, "crossAx", Text(crossingAxisId));
            axis.AppendVal(NsC, "crosses", "autoZero");
            axis.AppendVal(NsC, "crossBetween", "between");
            return axis;
        }

        private static OfficeXmlNode CreateDefaultDataLabels(OfficeXmlNode context) {
            OfficeXmlNode labels = context.CreateElement(NsC, "dLbls");
            labels.AppendVal(NsC, "showLegendKey", Text(false));
            labels.AppendVal(NsC, "showVal", Text(false));
            labels.AppendVal(NsC, "showCatName", Text(false));
            labels.AppendVal(NsC, "showSerName", Text(false));
            labels.AppendVal(NsC, "showPercent", Text(false));
            labels.AppendVal(NsC, "showBubbleSize", Text(false));
            return labels;
        }

        /// <summary>Creates a <c>c:strRef</c> with its cache for <paramref name="context"/>, the element it goes into.</summary>
        private static OfficeXmlNode CreateStringReference(OfficeXmlNode context, string formula, IReadOnlyList<string> values) {
            OfficeXmlNode reference = context.CreateElement(NsC, "strRef");
            FillStringReference(reference, formula, values);
            return reference;
        }

        /// <summary>Creates a <c>c:numRef</c> with its cache for <paramref name="context"/>, the element it goes into.</summary>
        private static OfficeXmlNode CreateNumberReference(OfficeXmlNode context, string formula, IReadOnlyList<double> values) {
            OfficeXmlNode reference = context.CreateElement(NsC, "numRef");
            FillNumberReference(reference, formula, values);
            return reference;
        }

        private static void AppendStringReference(OfficeXmlNode parent, string formula, IReadOnlyList<string> values) =>
            FillStringReference(parent.AppendNew(NsC, "strRef"), formula, values);

        private static void AppendNumberReference(OfficeXmlNode parent, string formula, IReadOnlyList<double> values) =>
            FillNumberReference(parent.AppendNew(NsC, "numRef"), formula, values);

        private static void FillStringReference(OfficeXmlNode reference, string formula, IReadOnlyList<string> values) {
            reference.AppendNew(NsC, "f").SetInnerText(formula);
            OfficeXmlNode cache = reference.AppendNew(NsC, "strCache");
            cache.AppendVal(NsC, "ptCount", Text((uint)values.Count));
            for (int i = 0; i < values.Count; i++) {
                OfficeXmlNode point = cache.AppendNew(NsC, "pt");
                point.SetAttribute("idx", Text((uint)i));
                point.AppendNew(NsC, "v").SetInnerText(values[i] ?? string.Empty);
            }
        }

        private static void FillNumberReference(OfficeXmlNode reference, string formula, IReadOnlyList<double> values) {
            reference.AppendNew(NsC, "f").SetInnerText(formula);
            OfficeXmlNode cache = reference.AppendNew(NsC, "numCache");
            cache.AppendNew(NsC, "formatCode").SetInnerText("General");
            cache.AppendVal(NsC, "ptCount", Text((uint)values.Count));
            for (int i = 0; i < values.Count; i++) {
                OfficeXmlNode point = cache.AppendNew(NsC, "pt");
                point.SetAttribute("idx", Text((uint)i));
                point.AppendNew(NsC, "v").SetInnerText(OfficeIMO.Core.Internal.OfficeRoundTripText.Format(values[i], CultureInfo.InvariantCulture));
            }
        }

        private static void UpdateSeriesIndexOrder(OfficeXmlNode series, int index) {
            OfficeXmlNode idx = series.Child(NsC, "idx");
            if (idx.IsNone) {
                OfficeXmlNode created = series.CreateElement(NsC, "idx");
                created.SetAttribute("val", Text((uint)index));
                idx = series.InsertBefore(created, series.FirstChild);
            } else {
                idx.SetAttribute("val", Text((uint)index));
            }

            OfficeXmlNode order = series.Child(NsC, "order");
            if (order.IsNone) {
                OfficeXmlNode created = series.CreateElement(NsC, "order");
                created.SetAttribute("val", Text((uint)index));
                idx.InsertAfterSelf(created);
            } else {
                order.SetAttribute("val", Text((uint)index));
            }
        }

        private static void UpdateScatterSeriesText(OfficeXmlNode series, int seriesIndex, string seriesName) {
            OfficeXmlNode text = series.CreateElement(NsC, "tx");
            AppendStringReference(text, GetScatterSeriesNameReference(seriesIndex), new[] { seriesName });
            SetOrdered(series, text, CScatterSerView.Order);
        }

        private static void UpdateXValues(OfficeXmlNode series, int seriesIndex, IReadOnlyList<double> values) {
            OfficeXmlNode xValues = series.CreateElement(NsC, "xVal");
            AppendNumberReference(xValues, GetScatterXValuesReference(seriesIndex, values.Count), values);
            SetOrdered(series, xValues, CScatterSerView.Order);
        }

        private static void UpdateYValues(OfficeXmlNode series, int seriesIndex, IReadOnlyList<double> values) {
            OfficeXmlNode yValues = series.CreateElement(NsC, "yVal");
            AppendNumberReference(yValues, GetScatterYValuesReference(seriesIndex, values.Count), values);
            SetOrdered(series, yValues, CScatterSerView.Order);
        }

        /// <summary>Inserts a series into a chart layer before its trailing chart-level elements and returns it as held.</summary>
        private static OfficeXmlNode InsertSeries(OfficeXmlNode chart, OfficeXmlNode series) {
            OfficeXmlNode insertBefore = default;
            for (OfficeXmlNode child = chart.FirstChild; !child.IsNone; child = child.NextSibling) {
                // A layer's c:marker is the show-marker flag, which the SDK reads as a different class than the series
                // marker, so it is not a stop here.
                if (child.NamespaceId == NsC && child.LocalName is "dLbls" or "gapWidth" or "overlap" or "axId" or
                    "firstSliceAng" or "holeSize" or "smooth" or "serLines") {
                    insertBefore = child;
                    break;
                }
            }

            return chart.InsertBefore(series, insertBefore);
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

        internal static void UpdateScatterData(IOfficeChartPart chartPart, OfficeChartData data) {
            ArgumentNullException.ThrowIfNull(chartPart);
            ArgumentNullException.ThrowIfNull(data);
            ValidateSharedChartData(data, OfficeChartType.Scatter);

            OfficeXmlNode plotArea = chartPart.ChartSpace.Child(NsC, "chart").Child(NsC, "plotArea");
            if (plotArea.IsNone) {
                throw new InvalidOperationException("Chart plot area not found.");
            }
            if (!IsOnlySharedScatterPlot(plotArea))
                throw new NotSupportedException("In-place scatter updates require a plot containing only scatter chart layers.");

            List<OfficeXmlNode> scatterCharts = plotArea.Children(NsC, "scatterChart").ToList();
            if (scatterCharts.Count > 0) {
                UpdateScatterChartLayers(plotArea, scatterCharts, data);
                return;
            }

            throw new NotSupportedException("Chart type is not supported for scatter data updates.");
        }

        // SDK callers
        internal static void UpdateScatterData(ChartPart chartPart, OfficeChartData data) =>
            UpdateScatterData(new OfficeSdkChartPart(chartPart ?? throw new ArgumentNullException(nameof(chartPart))), data);
    }
}
