using System;
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
        ///     Configures data labels for all bar chart series.
        /// </summary>
        public PowerPointChart SetDataLabels(bool showValue = true, bool showCategoryName = false,
            bool showSeriesName = false, bool showLegendKey = false, bool showPercent = false) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            foreach (C.BarChart barChart in plotArea.Elements<C.BarChart>()) {
                C.DataLabels labels = barChart.GetFirstChild<C.DataLabels>() ?? new C.DataLabels();
                ReplaceChild(labels, new C.ShowLegendKey { Val = showLegendKey });
                ReplaceChild(labels, new C.ShowValue { Val = showValue });
                ReplaceChild(labels, new C.ShowCategoryName { Val = showCategoryName });
                ReplaceChild(labels, new C.ShowSeriesName { Val = showSeriesName });
                ReplaceChild(labels, new C.ShowPercent { Val = showPercent });
                ReplaceChild(labels, new C.ShowBubbleSize { Val = false });

                if (barChart.GetFirstChild<C.DataLabels>() == null) {
                    barChart.Append(labels);
                }
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
