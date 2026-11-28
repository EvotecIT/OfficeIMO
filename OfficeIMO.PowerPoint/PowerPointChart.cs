using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
        /// <summary>
        ///     Represents a chart on a slide.
        /// </summary>
        public partial class PowerPointChart : PowerPointShape {
        internal PowerPointChart(PowerPointSlide slide, OfficeXmlTree tree, int node) : base(slide, tree, node) {
        }

        internal PowerPointChart(OpenXml.PowerPointPartRef part, OfficeXmlTree tree, int node) : base(part, tree, node) {
        }

        private OfficeTheme ColorTheme => PowerPointDrawingColor.ThemeFor(OwnerPartRef, OwnerSlide);

        /// <summary>
        ///     Updates the chart data using selectors.
        /// </summary>
        internal PowerPointChart UpdateData<T>(IEnumerable<T> items, Func<T, string> categorySelector,
            params PowerPointChartSeriesDefinition<T>[] seriesDefinitions) {
            PowerPointChartData data = PowerPointChartData.From(items, categorySelector, seriesDefinitions);
            return UpdateData(data);
        }

        /// <summary>
        ///     Updates scatter chart data using selectors.
        /// </summary>
        internal PowerPointChart UpdateData<T>(IEnumerable<T> items, Func<T, double> xSelector,
            params PowerPointScatterChartSeriesDefinition<T>[] seriesDefinitions) {
            PowerPointScatterChartData data = PowerPointScatterChartData.From(items, xSelector, seriesDefinitions);
            return UpdateData(data);
        }

        /// <summary>
        ///     Sets the chart title text.
        /// </summary>
        public PowerPointChart SetTitle(string title) {
            ArgumentNullException.ThrowIfNull(title);

            OfficeXmlNode chart = ChartNode;
            SetValChild(chart, "autoTitleDeleted", "0", CChartView.Order);

            OfficeXmlNode existing = chart.Child(CNs, "title");
            OfficeXmlNode chartTitle = existing.Exists ? existing : chart.CreateElement(CNs, "title");
            chartTitle.RemoveChildren(CNs, "tx");
            OfficeChartXml.SetOrdered(chartTitle, CreateChartText(chartTitle, title), CTitleView.Order);
            if (chartTitle.Child(CNs, "layout").IsNone) {
                chartTitle.AddOrdered(CNs, "layout", CTitleView.Order);
            }
            chartTitle.RemoveChildren(CNs, "overlay");
            OfficeChartXml.SetOrdered(chartTitle, chartTitle.CreateVal(CNs, "overlay", "0"), CTitleView.Order);

            if (existing.IsNone) {
                OfficeOpenXmlChartThemeWriter.ApplyTitle(chartTitle, ColorTheme);
                chart.InsertBefore(chartTitle, chart.FirstChild);
            }

            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Sets the chart title text style.
        /// </summary>
        public PowerPointChart SetTitleTextStyle(OfficeLength? fontSize = null, bool? bold = null, bool? italic = null,
            OfficeColor? color = null, string? fontFamily = null) {
            ValidateTextStyle(fontSize, color, fontFamily);

            OfficeXmlNode title = ChartNode.Child(CNs, "title");
            if (title.IsNone) {
                return this;
            }

            OfficeXmlNode chartText = title.Child(CNs, "tx");
            if (chartText.IsNone) {
                return this;
            }

            EnsureChartTextRunProperties(chartText);
            PowerPointTextContent.FormatRuns(chartText, properties => ApplyTextStyle(properties, fontSize, bold, italic, color, fontFamily));
            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Clears custom chart title text styling while preserving the title text.
        /// </summary>
        public PowerPointChart ClearTitleTextStyle() {
            OfficeXmlNode chartText = ChartNode.Child(CNs, "title").Child(CNs, "tx");
            if (chartText.IsNone) {
                return this;
            }

            foreach (OfficeXmlNode runProps in chartText.ContentDescendants(ANs, "rPr")) {
                ClearTextStyle(runProps);
            }

            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Removes the chart title.
        /// </summary>
        public PowerPointChart ClearTitle() {
            OfficeXmlNode chart = ChartNode;
            OfficeXmlNode title = chart.Child(CNs, "title");
            if (title.Exists) title.Remove();
            SetValChild(chart, "autoTitleDeleted", "1", CChartView.Order);
            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Sets the legend position and visibility.
        /// </summary>
        public PowerPointChart SetLegend(OfficeChartLegendPosition position, bool overlay = false) {
            OfficeXmlNode chart = ChartNode;
            OfficeXmlNode legend = chart.Child(CNs, "legend");
            if (legend.IsNone) {
                legend = chart.CreateElement(CNs, "legend");
                OfficeXmlNode plotArea = chart.Child(CNs, "plotArea");
                legend = plotArea.Exists ? plotArea.InsertAfterSelf(legend) : chart.Append(legend);
            }
            OfficeXmlNode legendPosition = legend.Child(CNs, "legendPos");
            if (legendPosition.IsNone) {
                legendPosition = legend.AddOrdered(CNs, "legendPos", CLegendView.Order);
            }
            legendPosition.SetAttribute("val", position.ToOpenXmlToken());
            if (legend.Child(CNs, "layout").IsNone) {
                legend.AddOrdered(CNs, "layout", CLegendView.Order);
            }
            legend.RemoveChildren(CNs, "overlay");
            InsertLegendOverlay(legend, legend.CreateElement(CNs, "overlay")).SetBooleanAttribute("val", overlay);

            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Sets the legend text style.
        /// </summary>
        public PowerPointChart SetLegendTextStyle(OfficeLength? fontSize = null, bool? bold = null, bool? italic = null,
            OfficeColor? color = null, string? fontFamily = null) {
            ValidateTextStyle(fontSize, color, fontFamily);

            OfficeXmlNode legend = ChartNode.Child(CNs, "legend");
            if (legend.IsNone) {
                return this;
            }

            ApplyTextStyle(EnsureTextPropertiesRunProperties(legend), fontSize, bold, italic, color, fontFamily);
            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Clears custom legend text styling while preserving the legend.
        /// </summary>
        public PowerPointChart ClearLegendTextStyle() {
            OfficeXmlNode runProps = ChartNode.Child(CNs, "legend").Child(CNs, "txPr").Child(ANs, "p").Child(ANs, "pPr").Child(ANs, "defRPr");
            if (runProps.IsNone) {
                return this;
            }

            ClearTextStyle(runProps);
            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Hides the chart legend.
        /// </summary>
        public PowerPointChart HideLegend() {
            OfficeXmlNode legend = ChartNode.Child(CNs, "legend");
            if (legend.Exists) legend.Remove();
            MarkChartEdited();
            return this;
        }

        private PowerPointChart ClearAxisLabelTextStyle(string axisName, Func<OfficeXmlNode, bool>? predicate = null) {
            OfficeXmlNode axis = FindAxis(axisName, predicate);
            OfficeXmlNode runProps = AxisLabelRunProperties(axis);
            if (runProps.IsNone) {
                return this;
            }

            ClearTextStyle(runProps);
            OfficeOpenXmlChartThemeWriter.RestoreAxisTextDefaults(runProps, ColorTheme);
            MarkChartEdited();
            return this;
        }

    }
}
