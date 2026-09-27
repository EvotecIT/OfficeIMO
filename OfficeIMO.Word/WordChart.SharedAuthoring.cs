using System.IO;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordChart {
    /// <summary>Gets or sets the native drawing name without changing chart data.</summary>
    public string? Name {
        get => _drawing?.Descendants<DW.DocProperties>().FirstOrDefault()?.Name?.Value;
        set {
            DW.DocProperties? properties = _drawing?.Descendants<DW.DocProperties>().FirstOrDefault();
            if (properties != null) properties.Name = value ?? string.Empty;
        }
    }

    /// <summary>Gets or sets alternative text persisted in the native Word drawing.</summary>
    public string? AltText {
        get => _drawing?.Descendants<DW.DocProperties>().FirstOrDefault()?.Description?.Value;
        set {
            DW.DocProperties? properties = _drawing?.Descendants<DW.DocProperties>().FirstOrDefault();
            if (properties != null) properties.Description = value;
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
    public WordChart SetData(OfficeChartKind chartKind, OfficeChartData data) =>
        ConfigureSharedData(chartKind, data, OfficeOpenXmlChartWriter.BuildWorkbook(data, chartKind));

    internal WordChart ConfigureSharedData(OfficeChartKind chartKind, OfficeChartData data, byte[] workbook) {
        ChartPart part = _chartPart ?? throw new InvalidOperationException("The chart has no native drawing part.");
        ChartSpace source = part.ChartSpace ?? throw new InvalidOperationException("The chart has no native chart space.");
        EmbeddedPackagePart embedded = OfficeOpenXmlChartWriter.GetSharedEmbeddedWorkbook(part) ??
            part.AddEmbeddedPackagePart(OfficeOpenXmlChartWriter.SharedWorkbookContentType);
        bool rounded = source.GetFirstChild<RoundedCorners>()?.Val?.Value ?? false;
        if (source.GetFirstChild<Chart>() == null) {
            OfficeOpenXmlChartWriter.PopulateSharedChart(part, part.GetIdOfPart(embedded), data, chartKind);
            part.ChartSpace!.GetFirstChild<RoundedCorners>()!.Val = rounded;
        } else {
            OfficeOpenXmlChartWriter.UpdateSharedChartData(part, data, chartKind);
            ExternalData? external = part.ChartSpace.GetFirstChild<ExternalData>();
            if (external == null) {
                external = new ExternalData();
                part.ChartSpace.AddChild(external, true);
            }
            external.Id = part.GetIdOfPart(embedded);
            external.AutoUpdate = new AutoUpdate { Val = false };
        }
        using (var stream = new MemoryStream(workbook, writable: false)) embedded.FeedData(stream);
        _chart = part.ChartSpace.GetFirstChild<Chart>();
        UpdateTitle();
        part.ChartSpace.Save();
        return this;
    }
}
