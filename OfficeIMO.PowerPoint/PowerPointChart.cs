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
        public partial class PowerPointChart : PowerPointShape {
        private readonly OpenXmlPartContainer _ownerPart;

        internal PowerPointChart(GraphicFrame frame, SlidePart slidePart) : base(frame) {
            _ownerPart = slidePart ?? throw new ArgumentNullException(nameof(slidePart));
        }

        internal PowerPointChart(GraphicFrame frame, OpenXmlPartContainer ownerPart) : base(frame) {
            _ownerPart = ownerPart ?? throw new ArgumentNullException(nameof(ownerPart));
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
        ///     Updates scatter chart data (series X/Y values).
        /// </summary>
        public PowerPointChart UpdateData(PowerPointScatterChartData data) {
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
        ///     Updates scatter chart data using selectors.
        /// </summary>
        public PowerPointChart UpdateData<T>(IEnumerable<T> items, Func<T, double> xSelector,
            params PowerPointScatterChartSeriesDefinition<T>[] seriesDefinitions) {
            PowerPointScatterChartData data = PowerPointScatterChartData.From(items, xSelector, seriesDefinitions);
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
        ///     Sets the chart title text style.
        /// </summary>
        public PowerPointChart SetTitleTextStyle(double? fontSizePoints = null, bool? bold = null, bool? italic = null,
            string? color = null, string? fontName = null) {
            ValidateTextStyle(fontSizePoints, color, fontName);

            C.Chart chart = GetChart();
            C.Title? title = chart.GetFirstChild<C.Title>();
            if (title == null) {
                return this;
            }

            C.ChartText? chartText = title.GetFirstChild<C.ChartText>();
            if (chartText == null) {
                return this;
            }

            ApplyTextStyle(EnsureChartTextRunProperties(chartText), fontSizePoints, bold, italic, color, fontName);
            Save();
            return this;
        }

        /// <summary>
        ///     Clears custom chart title text styling while preserving the title text.
        /// </summary>
        public PowerPointChart ClearTitleTextStyle() {
            C.Chart chart = GetChart();
            C.Title? title = chart.GetFirstChild<C.Title>();
            C.ChartText? chartText = title?.GetFirstChild<C.ChartText>();
            if (chartText == null) {
                return this;
            }

            foreach (A.RunProperties runProps in chartText.Descendants<A.RunProperties>()) {
                ClearTextStyle(runProps);
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
            InsertLegendOverlay(legend, new C.Overlay { Val = overlay });

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
        ///     Sets the legend text style.
        /// </summary>
        public PowerPointChart SetLegendTextStyle(double? fontSizePoints = null, bool? bold = null, bool? italic = null,
            string? color = null, string? fontName = null) {
            ValidateTextStyle(fontSizePoints, color, fontName);

            C.Chart chart = GetChart();
            C.Legend? legend = chart.GetFirstChild<C.Legend>();
            if (legend == null) {
                return this;
            }

            ApplyTextStyle(EnsureTextPropertiesRunProperties(legend), fontSizePoints, bold, italic, color, fontName);
            Save();
            return this;
        }

        /// <summary>
        ///     Clears custom legend text styling while preserving the legend.
        /// </summary>
        public PowerPointChart ClearLegendTextStyle() {
            C.Chart chart = GetChart();
            C.Legend? legend = chart.GetFirstChild<C.Legend>();
            A.DefaultRunProperties? runProps = legend?
                .GetFirstChild<C.TextProperties>()?
                .GetFirstChild<A.Paragraph>()?
                .GetFirstChild<A.ParagraphProperties>()?
                .GetFirstChild<A.DefaultRunProperties>();
            if (runProps == null) {
                return this;
            }

            ClearTextStyle(runProps);
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
            return (ChartPart)_ownerPart.GetPartById(relId);
        }

        private void Save() {
            ChartPart chartPart = GetChartPart();
            chartPart.ChartSpace?.Save();
        }
    }
}
