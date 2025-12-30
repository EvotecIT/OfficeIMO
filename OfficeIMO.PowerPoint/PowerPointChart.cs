using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
        /// <summary>
        ///     Represents a chart on a slide.
        /// </summary>
        public class PowerPointChart : PowerPointShape {
        private readonly SlidePart _slidePart;

        internal PowerPointChart(GraphicFrame frame, SlidePart slidePart) : base(frame) {
            _slidePart = slidePart ?? throw new ArgumentNullException(nameof(slidePart));
        }

        private GraphicFrame Frame => (GraphicFrame)Element;

        /// <summary>
        ///     Updates the chart data (series and categories).
        /// </summary>
        public PowerPointChart UpdateData(PowerPointChartData data) {
            if (data == null) {
                throw new ArgumentNullException(nameof(data));
            }

            ChartPart chartPart = GetChartPart();
            PowerPointUtils.UpdateChartData(chartPart, data);

            EmbeddedPackagePart? embedded = chartPart.GetPartsOfType<EmbeddedPackagePart>().FirstOrDefault();
            if (embedded != null) {
                byte[] workbookBytes = PowerPointUtils.BuildChartWorkbook(data);
                using var stream = new MemoryStream(workbookBytes);
                embedded.FeedData(stream);
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Updates the chart data using selectors.
        /// </summary>
        public PowerPointChart UpdateData<T>(IEnumerable<T> items, Func<T, string> categorySelector,
            params PowerPointChartSeriesDefinition<T>[] seriesDefinitions) {
            PowerPointChartData data = PowerPointChartData.From(items, categorySelector, seriesDefinitions);
            return UpdateData(data);
        }

        /// <summary>
        ///     Sets the chart title text.
        /// </summary>
        public PowerPointChart SetTitle(string title) {
            if (title == null) {
                throw new ArgumentNullException(nameof(title));
            }

            C.Chart chart = GetChart();
            chart.AutoTitleDeleted = new C.AutoTitleDeleted { Val = false };

            C.Title chartTitle = chart.GetFirstChild<C.Title>() ?? new C.Title();
            chartTitle.RemoveAllChildren<C.ChartText>();
            chartTitle.Append(CreateChartText(title));
            if (chartTitle.GetFirstChild<C.Layout>() == null) {
                chartTitle.Append(new C.Layout());
            }
            chartTitle.RemoveAllChildren<C.Overlay>();
            chartTitle.Append(new C.Overlay { Val = false });

            if (chart.GetFirstChild<C.Title>() == null) {
                chart.InsertAt(chartTitle, 0);
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Removes the chart title.
        /// </summary>
        public PowerPointChart ClearTitle() {
            C.Chart chart = GetChart();
            chart.GetFirstChild<C.Title>()?.Remove();
            chart.AutoTitleDeleted = new C.AutoTitleDeleted { Val = true };
            Save();
            return this;
        }

        /// <summary>
        ///     Sets the legend position and visibility.
        /// </summary>
        public PowerPointChart SetLegend(C.LegendPositionValues position, bool overlay = false) {
            C.Chart chart = GetChart();
            C.Legend legend = chart.GetFirstChild<C.Legend>() ?? new C.Legend();
            var legendPosition = legend.GetFirstChild<C.LegendPosition>() ?? new C.LegendPosition();
            legendPosition.Val = position;
            if (legendPosition.Parent == null) {
                legend.Append(legendPosition);
            }
            if (legend.GetFirstChild<C.Layout>() == null) {
                legend.Append(new C.Layout());
            }
            legend.RemoveAllChildren<C.Overlay>();
            legend.Append(new C.Overlay { Val = overlay });

            if (chart.GetFirstChild<C.Legend>() == null) {
                C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
                if (plotArea != null) {
                    chart.InsertAfter(legend, plotArea);
                } else {
                    chart.Append(legend);
                }
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Hides the chart legend.
        /// </summary>
        public PowerPointChart HideLegend() {
            C.Chart chart = GetChart();
            chart.GetFirstChild<C.Legend>()?.Remove();
            Save();
            return this;
        }

        /// <summary>
        ///     Configures data labels for all supported chart series.
        /// </summary>
        public PowerPointChart SetDataLabels(bool showValue = true, bool showCategoryName = false,
            bool showSeriesName = false, bool showLegendKey = false, bool showPercent = false) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            foreach (C.BarChart barChart in plotArea.Elements<C.BarChart>()) {
                ApplyDataLabels(barChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent);
            }

            foreach (C.LineChart lineChart in plotArea.Elements<C.LineChart>()) {
                ApplyDataLabels(lineChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent);
            }

            foreach (C.AreaChart areaChart in plotArea.Elements<C.AreaChart>()) {
                ApplyDataLabels(areaChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent);
            }

            foreach (C.PieChart pieChart in plotArea.Elements<C.PieChart>()) {
                ApplyDataLabels(pieChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent);
            }

            foreach (C.DoughnutChart doughnutChart in plotArea.Elements<C.DoughnutChart>()) {
                ApplyDataLabels(doughnutChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent);
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets the category axis title.
        /// </summary>
        public PowerPointChart SetCategoryAxisTitle(string title) {
            return SetAxisTitle<C.CategoryAxis>(title);
        }

        /// <summary>
        ///     Sets the value axis title.
        /// </summary>
        public PowerPointChart SetValueAxisTitle(string title) {
            return SetAxisTitle<C.ValueAxis>(title);
        }

        /// <summary>
        ///     Sets the category axis number format.
        /// </summary>
        public PowerPointChart SetCategoryAxisNumberFormat(string formatCode, bool sourceLinked = false) {
            return SetAxisNumberFormat<C.CategoryAxis>(formatCode, sourceLinked);
        }

        /// <summary>
        ///     Sets the value axis number format.
        /// </summary>
        public PowerPointChart SetValueAxisNumberFormat(string formatCode, bool sourceLinked = false) {
            return SetAxisNumberFormat<C.ValueAxis>(formatCode, sourceLinked);
        }

        /// <summary>
        ///     Sets the fill color for a chart series by index.
        /// </summary>
        public PowerPointChart SetSeriesFillColor(int seriesIndex, string color) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Series color cannot be null or empty.", nameof(color));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                C.ChartShapeProperties props = EnsureChartShapeProperties(series);
                ApplySolidFill(props, color);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets the fill color for a chart series by name.
        /// </summary>
        public PowerPointChart SetSeriesFillColor(string seriesName, string color, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Series color cannot be null or empty.", nameof(color));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                C.ChartShapeProperties props = EnsureChartShapeProperties(series);
                ApplySolidFill(props, color);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets the line color for a chart series by index.
        /// </summary>
        public PowerPointChart SetSeriesLineColor(int seriesIndex, string color, double? widthPoints = null) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Series color cannot be null or empty.", nameof(color));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                C.ChartShapeProperties props = EnsureChartShapeProperties(series);
                ApplyLine(props, color, widthPoints);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets the line color for a chart series by name.
        /// </summary>
        public PowerPointChart SetSeriesLineColor(string seriesName, string color, double? widthPoints = null, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Series color cannot be null or empty.", nameof(color));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                C.ChartShapeProperties props = EnsureChartShapeProperties(series);
                ApplyLine(props, color, widthPoints);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets the marker style for a chart series by index.
        /// </summary>
        public PowerPointChart SetSeriesMarker(int seriesIndex, C.MarkerStyleValues style, int? size = null, string? fillColor = null, string? lineColor = null, double? lineWidthPoints = null) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (size is < 1 or > 72) {
                throw new ArgumentOutOfRangeException(nameof(size), "Marker size must be between 1 and 72.");
            }
            if (fillColor != null && string.IsNullOrWhiteSpace(fillColor)) {
                throw new ArgumentException("Marker fill color cannot be empty.", nameof(fillColor));
            }
            if (lineColor != null && string.IsNullOrWhiteSpace(lineColor)) {
                throw new ArgumentException("Marker line color cannot be empty.", nameof(lineColor));
            }

            bool applied = ApplySeriesMarkerByIndex(seriesIndex, marker => {
                ApplyMarker(marker, style, size, fillColor, lineColor, lineWidthPoints);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets the marker style for a chart series by name.
        /// </summary>
        public PowerPointChart SetSeriesMarker(string seriesName, C.MarkerStyleValues style, int? size = null, string? fillColor = null, string? lineColor = null, double? lineWidthPoints = null, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (size is < 1 or > 72) {
                throw new ArgumentOutOfRangeException(nameof(size), "Marker size must be between 1 and 72.");
            }
            if (fillColor != null && string.IsNullOrWhiteSpace(fillColor)) {
                throw new ArgumentException("Marker fill color cannot be empty.", nameof(fillColor));
            }
            if (lineColor != null && string.IsNullOrWhiteSpace(lineColor)) {
                throw new ArgumentException("Marker line color cannot be empty.", nameof(lineColor));
            }

            bool applied = ApplySeriesMarkerByName(seriesName, ignoreCase, marker => {
                ApplyMarker(marker, style, size, fillColor, lineColor, lineWidthPoints);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        private PowerPointChart SetAxisTitle<TAxis>(string title) where TAxis : OpenXmlCompositeElement {
            if (title == null) {
                throw new ArgumentNullException(nameof(title));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            TAxis? axis = plotArea.Elements<TAxis>().FirstOrDefault();
            if (axis == null) {
                return this;
            }

            axis.RemoveAllChildren<C.Title>();
            axis.Append(CreateAxisTitle(title));
            Save();
            return this;
        }

        private PowerPointChart SetAxisNumberFormat<TAxis>(string formatCode, bool sourceLinked)
            where TAxis : OpenXmlCompositeElement {
            if (string.IsNullOrWhiteSpace(formatCode)) {
                throw new ArgumentException("Format code cannot be null or empty.", nameof(formatCode));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            TAxis? axis = plotArea.Elements<TAxis>().FirstOrDefault();
            if (axis == null) {
                return this;
            }

            C.NumberingFormat format = axis.GetFirstChild<C.NumberingFormat>() ?? new C.NumberingFormat();
            format.FormatCode = formatCode;
            format.SourceLinked = sourceLinked;
            if (format.Parent == null) {
                axis.InsertAt(format, 0);
            }

            Save();
            return this;
        }

        private static void ApplyDataLabels(OpenXmlCompositeElement chartElement, bool showLegendKey, bool showValue,
            bool showCategoryName, bool showSeriesName, bool showPercent) {
            C.DataLabels labels = chartElement.GetFirstChild<C.DataLabels>() ?? new C.DataLabels();
            ReplaceChild(labels, new C.ShowLegendKey { Val = showLegendKey });
            ReplaceChild(labels, new C.ShowValue { Val = showValue });
            ReplaceChild(labels, new C.ShowCategoryName { Val = showCategoryName });
            ReplaceChild(labels, new C.ShowSeriesName { Val = showSeriesName });
            ReplaceChild(labels, new C.ShowPercent { Val = showPercent });
            ReplaceChild(labels, new C.ShowBubbleSize { Val = false });

            if (chartElement.GetFirstChild<C.DataLabels>() == null) {
                chartElement.Append(labels);
            }
        }

        private static C.ChartShapeProperties EnsureChartShapeProperties(OpenXmlCompositeElement series) {
            C.ChartShapeProperties props = series.GetFirstChild<C.ChartShapeProperties>() ?? new C.ChartShapeProperties();
            if (props.Parent == null) {
                series.Append(props);
            }
            return props;
        }

        private static void ApplySolidFill(OpenXmlCompositeElement props, string color) {
            props.RemoveAllChildren<A.SolidFill>();
            props.RemoveAllChildren<A.NoFill>();
            props.RemoveAllChildren<A.GradientFill>();
            props.RemoveAllChildren<A.PatternFill>();
            props.Append(new A.SolidFill(new A.RgbColorModelHex { Val = color }));
        }

        private static void ApplyLine(OpenXmlCompositeElement props, string color, double? widthPoints) {
            A.Outline outline = props.GetFirstChild<A.Outline>() ?? new A.Outline();
            outline.RemoveAllChildren<A.SolidFill>();
            outline.Append(new A.SolidFill(new A.RgbColorModelHex { Val = color }));
            if (widthPoints != null) {
                outline.Width = (int)Math.Round(widthPoints.Value * 12700d);
            }

            if (outline.Parent == null) {
                props.Append(outline);
            }
        }

        private static void ApplyOptionalLine(OpenXmlCompositeElement props, string? color, double? widthPoints) {
            if (color == null && widthPoints == null) {
                return;
            }

            A.Outline outline = props.GetFirstChild<A.Outline>() ?? new A.Outline();
            if (color != null) {
                outline.RemoveAllChildren<A.SolidFill>();
                outline.Append(new A.SolidFill(new A.RgbColorModelHex { Val = color }));
            }
            if (widthPoints != null) {
                outline.Width = (int)Math.Round(widthPoints.Value * 12700d);
            }

            if (outline.Parent == null) {
                props.Append(outline);
            }
        }

        private static void ApplyMarker(C.Marker marker, C.MarkerStyleValues style, int? size, string? fillColor, string? lineColor, double? lineWidthPoints) {
            marker.Symbol = new C.Symbol { Val = style };
            if (size != null) {
                marker.Size = new C.Size { Val = (byte)size.Value };
            }

            if (fillColor != null || lineColor != null || lineWidthPoints != null) {
                C.ChartShapeProperties props = marker.ChartShapeProperties ?? new C.ChartShapeProperties();
                if (fillColor != null) {
                    ApplySolidFill(props, fillColor);
                }
                if (lineColor != null || lineWidthPoints != null) {
                    ApplyOptionalLine(props, lineColor, lineWidthPoints);
                }
                if (props.Parent == null) {
                    marker.Append(props);
                }
            }
        }

        private bool ApplySeriesByIndex(int seriesIndex, Action<OpenXmlCompositeElement> apply) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return false;
            }

            if (ApplySeriesByIndex(plotArea.Elements<C.BarChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.LineChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.AreaChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.PieChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.DoughnutChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.ScatterChart>(), seriesIndex, apply)) return true;

            return false;
        }

        private bool ApplySeriesByName(string seriesName, bool ignoreCase, Action<OpenXmlCompositeElement> apply) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return false;
            }

            if (ApplySeriesByName(plotArea.Elements<C.BarChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.LineChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.AreaChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.PieChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.DoughnutChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.ScatterChart>(), seriesName, ignoreCase, apply)) return true;

            return false;
        }

        private bool ApplySeriesMarkerByIndex(int seriesIndex, Action<C.Marker> apply) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return false;
            }

            if (ApplySeriesMarkerByIndex(plotArea.Elements<C.LineChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesMarkerByIndex(plotArea.Elements<C.ScatterChart>(), seriesIndex, apply)) return true;

            return false;
        }

        private bool ApplySeriesMarkerByName(string seriesName, bool ignoreCase, Action<C.Marker> apply) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return false;
            }

            if (ApplySeriesMarkerByName(plotArea.Elements<C.LineChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesMarkerByName(plotArea.Elements<C.ScatterChart>(), seriesName, ignoreCase, apply)) return true;

            return false;
        }

        private static bool ApplySeriesMarkerByIndex<TChart>(IEnumerable<TChart> charts, int seriesIndex, Action<C.Marker> apply) where TChart : OpenXmlCompositeElement {
            foreach (TChart chart in charts) {
                List<OpenXmlCompositeElement> series = chart.ChildElements
                    .OfType<OpenXmlCompositeElement>()
                    .Where(IsSeriesElement)
                    .OrderBy(GetSeriesIndex)
                    .ToList();

                if (seriesIndex < 0 || seriesIndex >= series.Count) {
                    continue;
                }

                OpenXmlCompositeElement seriesElement = series[seriesIndex];
                C.Marker marker = seriesElement.GetFirstChild<C.Marker>() ?? new C.Marker();
                apply(marker);
                if (marker.Parent == null) {
                    seriesElement.Append(marker);
                }
                return true;
            }

            return false;
        }

        private static bool ApplySeriesMarkerByName<TChart>(IEnumerable<TChart> charts, string seriesName, bool ignoreCase, Action<C.Marker> apply) where TChart : OpenXmlCompositeElement {
            StringComparison comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            foreach (TChart chart in charts) {
                foreach (OpenXmlCompositeElement series in chart.ChildElements.OfType<OpenXmlCompositeElement>().Where(IsSeriesElement)) {
                    string? name = GetSeriesName(series);
                    if (name != null && string.Equals(name, seriesName, comparison)) {
                        C.Marker marker = series.GetFirstChild<C.Marker>() ?? new C.Marker();
                        apply(marker);
                        if (marker.Parent == null) {
                            series.Append(marker);
                        }
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool ApplySeriesByIndex<TChart>(IEnumerable<TChart> charts, int seriesIndex,
            Action<OpenXmlCompositeElement> apply) where TChart : OpenXmlCompositeElement {
            foreach (TChart chart in charts) {
                List<OpenXmlCompositeElement> series = chart.ChildElements
                    .OfType<OpenXmlCompositeElement>()
                    .Where(IsSeriesElement)
                    .OrderBy(GetSeriesIndex)
                    .ToList();

                if (seriesIndex < 0 || seriesIndex >= series.Count) {
                    continue;
                }

                apply(series[seriesIndex]);
                return true;
            }

            return false;
        }

        private static bool ApplySeriesByName<TChart>(IEnumerable<TChart> charts, string seriesName, bool ignoreCase,
            Action<OpenXmlCompositeElement> apply) where TChart : OpenXmlCompositeElement {
            StringComparison comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            foreach (TChart chart in charts) {
                foreach (OpenXmlCompositeElement series in chart.ChildElements.OfType<OpenXmlCompositeElement>().Where(IsSeriesElement)) {
                    string? name = GetSeriesName(series);
                    if (name != null && string.Equals(name, seriesName, comparison)) {
                        apply(series);
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool IsSeriesElement(OpenXmlCompositeElement element) {
            return element is C.BarChartSeries ||
                   element is C.LineChartSeries ||
                   element is C.AreaChartSeries ||
                   element is C.PieChartSeries ||
                   element is C.ScatterChartSeries;
        }

        private static int GetSeriesIndex(OpenXmlCompositeElement series) {
            return (int)(series.GetFirstChild<C.Index>()?.Val?.Value ?? 0U);
        }

        private static string? GetSeriesName(OpenXmlCompositeElement series) {
            C.SeriesText? seriesText = series.GetFirstChild<C.SeriesText>();
            if (seriesText == null) {
                return null;
            }

            C.StringReference? reference = seriesText.GetFirstChild<C.StringReference>();
            C.StringCache? cache = reference?.GetFirstChild<C.StringCache>();
            string? cachedText = cache?.Elements<C.StringPoint>()
                .FirstOrDefault()?
                .NumericValue?
                .Text;
            if (!string.IsNullOrWhiteSpace(cachedText)) {
                return cachedText;
            }

            C.StringLiteral? literal = seriesText.GetFirstChild<C.StringLiteral>();
            string? literalText = literal?.Elements<C.StringPoint>()
                .FirstOrDefault()?
                .NumericValue?
                .Text;
            if (!string.IsNullOrWhiteSpace(literalText)) {
                return literalText;
            }

            return string.IsNullOrWhiteSpace(seriesText.InnerText) ? null : seriesText.InnerText;
        }

        private static C.ChartText CreateChartText(string title) {
            return new C.ChartText(
                new C.RichText(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(
                        new A.Run(
                            new A.RunProperties { Language = "en-US" },
                            new A.Text { Text = title })
                    )));
        }

        private static C.Title CreateAxisTitle(string title) {
            return new C.Title(
                new C.ChartText(
                    new C.RichText(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(
                            new A.Run(
                                new A.RunProperties { Language = "en-US" },
                                new A.Text { Text = title })))
                ),
                new C.Layout(),
                new C.Overlay { Val = false }
            );
        }

        private static void ReplaceChild<T>(OpenXmlCompositeElement parent, T child) where T : OpenXmlElement {
            parent.GetFirstChild<T>()?.Remove();
            parent.Append(child);
        }

        private C.Chart GetChart() {
            ChartPart chartPart = GetChartPart();
            C.Chart? chart = chartPart.ChartSpace?.GetFirstChild<C.Chart>();
            if (chart == null) {
                throw new InvalidOperationException("Chart element not found in chart part.");
            }
            return chart;
        }

        private ChartPart GetChartPart() {
            C.ChartReference? chartReference = Frame.Graphic?.GraphicData?.GetFirstChild<C.ChartReference>();
            StringValue? relationshipId = chartReference?.Id;
            if (relationshipId == null) {
                throw new InvalidOperationException("Chart reference not found for the shape.");
            }

            string relId = relationshipId.Value ?? throw new InvalidOperationException("Chart relationship id is empty.");
            return (ChartPart)_slidePart.GetPartById(relId);
        }

        private void Save() {
            ChartPart chartPart = GetChartPart();
            chartPart.ChartSpace?.Save();
        }
    }
}
