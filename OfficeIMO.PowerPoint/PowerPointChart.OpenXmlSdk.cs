using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// The chart members that still work on SDK types: the constructors chart creation calls with the SDK frame and
    /// part, and the internal data updates that go through the SDK chart builder.
    /// </summary>
    public partial class PowerPointChart {
        private readonly OpenXmlPartContainer? _initialOwnerPart;
        // A chart read from the package belongs to its slide's part once the SDK opens.
        private OpenXmlPartContainer _ownerPart => _initialOwnerPart is null or SlidePart
            ? Element.Ancestors<Slide>().FirstOrDefault()?.SlidePart ?? _initialOwnerPart ?? RelationshipOwnerPart ?? OwnerSlide!.SlidePart
            : _initialOwnerPart;

        internal PowerPointChart(GraphicFrame frame, SlidePart slidePart) : base(frame) {
            _initialOwnerPart = slidePart ?? throw new ArgumentNullException(nameof(slidePart));
        }

        internal PowerPointChart(GraphicFrame frame, OpenXmlPartContainer ownerPart) : base(frame) {
            _initialOwnerPart = ownerPart ?? throw new ArgumentNullException(nameof(ownerPart));
        }

        private GraphicFrame Frame => (GraphicFrame)Element;

        /// <summary>
        ///     Updates the chart data (series and categories).
        /// </summary>
        internal PowerPointChart UpdateData(PowerPointChartData data) {
            ArgumentNullException.ThrowIfNull(data);

            ChartPart chartPart = GetChartPart();
            OfficeIMO.OpenXml.Internal.IOfficeChartPart sharedPart = SharedChartPart;
            OfficeIMO.OpenXml.Internal.IOfficeChartEmbeddedPackage? embedded = OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartWriter.GetSharedEmbeddedWorkbook(sharedPart);
            byte[]? workbookBytes = embedded != null ? PowerPointUtils.BuildChartWorkbook(data) : null;
            Action? preserveBindings = embedded != null ? OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartWriter.PrepareSharedWorkbookBindings(
                sharedPart, data.Series.Select(series => series.Values.Count).ToArray()) : null;
            PowerPointUtils.UpdateChartData(chartPart, data);
            preserveBindings?.Invoke();
            embedded?.Replace(workbookBytes!);

            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Updates scatter chart data (series X/Y values).
        /// </summary>
        internal PowerPointChart UpdateData(PowerPointScatterChartData data) {
            ArgumentNullException.ThrowIfNull(data);

            ChartPart chartPart = GetChartPart();
            OfficeIMO.OpenXml.Internal.IOfficeChartPart sharedPart = SharedChartPart;
            OfficeIMO.OpenXml.Internal.IOfficeChartEmbeddedPackage? embedded = OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartWriter.GetSharedEmbeddedWorkbook(sharedPart);
            byte[]? workbookBytes = embedded != null ? PowerPointUtils.BuildChartWorkbook(data) : null;
            Action? preserveBindings = embedded != null ? OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartWriter.PrepareSharedWorkbookBindings(
                sharedPart, data.Series.Select(series => series.YValues.Count).ToArray()) : null;
            PowerPointUtils.UpdateChartData(chartPart, data);
            preserveBindings?.Invoke();
            embedded?.Replace(workbookBytes!);

            MarkChartEdited();
            return this;
        }

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
