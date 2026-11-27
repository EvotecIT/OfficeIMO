using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
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
        private readonly OpenXmlPartContainer? _initialOwnerPart;
        // A chart read from the package belongs to its slide's part once the SDK opens.
        private OpenXmlPartContainer _ownerPart => _initialOwnerPart is null or SlidePart
            ? Element.Ancestors<Slide>().FirstOrDefault()?.SlidePart ?? _initialOwnerPart ?? OwnerSlide!.SlidePart
            : _initialOwnerPart;

        internal PowerPointChart(PowerPointSlide slide, OfficeXmlTree tree, int node) : base(slide, tree, node) {
        }

        internal PowerPointChart(GraphicFrame frame, SlidePart slidePart) : base(frame) {
            _initialOwnerPart = slidePart ?? throw new ArgumentNullException(nameof(slidePart));
        }

        internal PowerPointChart(GraphicFrame frame, OpenXmlPartContainer ownerPart) : base(frame) {
            _initialOwnerPart = ownerPart ?? throw new ArgumentNullException(nameof(ownerPart));
        }

        private OfficeTheme ColorTheme => PowerPointDrawingColor.ThemeFor(OwnerPartRef, OwnerSlide);

        private GraphicFrame Frame => (GraphicFrame)Element;

        /// <summary>
        ///     Updates the chart data (series and categories).
        /// </summary>
        internal PowerPointChart UpdateData(PowerPointChartData data) {
            ArgumentNullException.ThrowIfNull(data);

            ChartPart chartPart = GetChartPart();
            EmbeddedPackagePart? embedded = OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartWriter.GetSharedEmbeddedWorkbook(chartPart);
            byte[]? workbookBytes = embedded != null ? PowerPointUtils.BuildChartWorkbook(data) : null;
            Action? preserveBindings = embedded != null ? OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartWriter.PrepareSharedWorkbookBindings(
                chartPart, data.Series.Select(series => series.Values.Count).ToArray()) : null;
            PowerPointUtils.UpdateChartData(chartPart, data);
            preserveBindings?.Invoke();
            if (embedded != null) {
                using var stream = new MemoryStream(workbookBytes!);
                embedded.FeedData(stream);
            }

            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Updates scatter chart data (series X/Y values).
        /// </summary>
        internal PowerPointChart UpdateData(PowerPointScatterChartData data) {
            ArgumentNullException.ThrowIfNull(data);

            ChartPart chartPart = GetChartPart();
            EmbeddedPackagePart? embedded = OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartWriter.GetSharedEmbeddedWorkbook(chartPart);
            byte[]? workbookBytes = embedded != null ? PowerPointUtils.BuildChartWorkbook(data) : null;
            Action? preserveBindings = embedded != null ? OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartWriter.PrepareSharedWorkbookBindings(
                chartPart, data.Series.Select(series => series.YValues.Count).ToArray()) : null;
            PowerPointUtils.UpdateChartData(chartPart, data);
            preserveBindings?.Invoke();
            if (embedded != null) {
                using var stream = new MemoryStream(workbookBytes!);
                embedded.FeedData(stream);
            }

            MarkChartEdited();
            return this;
        }

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
            if (chart.Child(CNs, "title").IsNone && !chart.IsSdk) {
                // A new title takes the theme's title text defaults through the shared writer, which works on SDK elements.
                RequireSdkChart();
                chart = ChartNode;
            }
            SetValChild(chart, "autoTitleDeleted", "0", CChartView.Order);

            OfficeXmlNode existing = chart.Child(CNs, "title");
            OfficeXmlNode chartTitle = existing.Exists ? existing : chart.CreateElement(CNs, "title");
            chartTitle.RemoveChildren(CNs, "tx");
            chartTitle.InsertOrdered(CreateChartText(chartTitle, title), CTitleView.Order);
            if (chartTitle.Child(CNs, "layout").IsNone) {
                chartTitle.AddOrdered(CNs, "layout", CTitleView.Order);
            }
            chartTitle.RemoveChildren(CNs, "overlay");
            chartTitle.AddOrdered(CNs, "overlay", CTitleView.Order).SetBooleanAttribute("val", false);

            if (existing.IsNone) {
                OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartThemeWriter.ApplyTitle((C.Title)chartTitle.SdkElement!,
                    ColorTheme);
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

            foreach (OfficeXmlNode runProps in chartText.Descendants(ANs, "rPr")) {
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
            if (axis.IsNone || AxisLabelRunProperties(axis).IsNone) {
                return this;
            }
            if (!axis.IsSdk) {
                // The theme's axis text defaults go back through the shared writer, which works on SDK elements.
                RequireSdkChart();
                axis = FindAxis(axisName, predicate);
            }

            OfficeXmlNode runProps = AxisLabelRunProperties(axis);
            ClearTextStyle(runProps);
            OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartThemeWriter.RestoreAxisTextDefaults((A.TextCharacterPropertiesType)runProps.SdkElement!,
                ColorTheme);
            MarkChartEdited();
            return this;
        }

        // Clears the automatic color marker of a series through the shared writer, which works on SDK elements; callers
        // open the SDK for the chart first.
        private static void ClearAutomaticSeriesColor(OfficeXmlNode series) =>
            OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartAutomaticColors.ClearAutomaticSeriesColor((OpenXmlElement)series.SdkElement!);

        /// <summary>Opens the SDK for the chart, so its elements are the SDK's.</summary>
        private void RequireSdkChart() => _ = GetChartPart();

        internal C.Chart GetChart() {
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
    }
}
