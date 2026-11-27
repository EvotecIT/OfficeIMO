using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word;

public partial class WordChart {
    /// <summary>The drawing's <c>wp:docPr</c>, or none.</summary>
    private OfficeXmlNode DrawingProperties {
        get {
            if (_drawing.IsNone) return default;
            List<OfficeXmlNode> properties = _drawing.ContentDescendants(NsWp, "docPr");
            return properties.Count == 0 ? default : properties[0];
        }
    }

    /// <summary>Gets or sets the native drawing name without changing chart data.</summary>
    public string? Name {
        get => DrawingProperties.IsNone ? null : DrawingProperties.Attribute("name");
        set {
            OfficeXmlNode properties = DrawingProperties;
            if (!properties.IsNone) properties.SetAttribute("name", value ?? string.Empty);
        }
    }

    /// <summary>Gets or sets alternative text persisted in the native Word drawing.</summary>
    public string? AltText {
        get => DrawingProperties.IsNone ? null : DrawingProperties.Attribute("descr");
        set {
            OfficeXmlNode properties = DrawingProperties;
            if (!properties.IsNone) properties.SetAttribute("descr", value);
        }
    }

    /// <summary>
    /// Initializes or updates an editable native chart and its embedded worksheet from shared chart data.
    /// Existing chart and drawing titles, dimensions, and accessibility text are preserved.
    /// Use series render kinds and axis groups to describe combinations and secondary value axes.
    /// </summary>
    /// <param name="chartKind">The default chart family for series without their own render kind.</param>
    /// <param name="data">Categories and series, including supported point and marker styles.</param>
    /// <returns>The current chart.</returns>
    public WordChart SetData(OfficeChartType chartKind, OfficeChartData data) =>
        ConfigureSharedData(chartKind, data, OfficeOpenXmlChartWriter.BuildWorkbook(data, chartKind));

    internal WordChart ConfigureSharedData(OfficeChartType chartKind, OfficeChartData data, byte[] workbook) {
        IOfficeChartPart part = SharedChartPart ?? throw new InvalidOperationException("The chart has no native drawing part.");
        OfficeXmlNode source = part.ChartSpace;
        if (source.IsNone) throw new InvalidOperationException("The chart has no native chart space.");
        var preserveBindings = OfficeOpenXmlChartWriter.PrepareSharedWorkbookBindings(
            part, data.Series.Select(series => series.Values.Count).ToArray());
        string? embedded = OfficeOpenXmlChartWriter.GetSharedEmbeddedWorkbook(part)?.RelationshipId;
        OfficeXmlNode roundedCorners = source.Child(NsC, "roundedCorners");
        bool rounded = (roundedCorners.IsNone ? null : roundedCorners.BoolVal()) ?? false;
        bool written = false;
        if (source.Child(NsC, "chart").IsNone) {
            embedded ??= AddEmbeddedWorkbook();
            // The workbook goes in before the chart that references it is saved to its part.
            WriteEmbeddedPackage(embedded, workbook);
            written = true;
            ReplaceChartSpaceRoot();
            OfficeOpenXmlChartWriter.PopulateSharedChart(part, embedded, data, chartKind);
            OfficeOpenXmlChartThemeWriter.Apply(part, OfficeTheme.Default);
            ChartSpaceNode.Child(NsC, "roundedCorners").SetAttribute("val", OfficeChartXml.Text(rounded));
            EnsureFramelessChartArea();
        } else {
            OfficeOpenXmlChartWriter.UpdateSharedChartData(part, data, chartKind);
            embedded ??= AddEmbeddedWorkbook();
            OfficeXmlNode space = ChartSpaceNode;
            OfficeXmlNode external = space.Child(NsC, "externalData");
            if (external.IsNone) external = OfficeChartXml.SetOrdered(space, NewC(space, "externalData"), CChartSpaceView.Order);
            external.SetAttribute(NsR, "id", embedded);
            OfficeChartXml.SetOrdered(external, NewC(external, "autoUpdate", false), CExternalDataView.Order);
        }
        preserveBindings();
        if (!written) WriteEmbeddedPackage(embedded, workbook);
        _chart = ChartSpaceNode.Child(NsC, "chart");
        UpdateTitle();
        SaveChartSpace();
        return this;
    }
}
