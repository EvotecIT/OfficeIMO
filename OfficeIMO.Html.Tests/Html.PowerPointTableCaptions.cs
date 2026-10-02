using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using Xunit;
namespace OfficeIMO.Tests;
public partial class HtmlOfficeAdapters {
    private const string RegulatoryTableCaption = "List of National Secondary Drinking Water Regulations";
    private const string RegulatoryTableLink = "https://example.org/standards";
    private const string RegulatoryTableHtml = "<main><h1>Drinking water</h1><table><caption><a href='"
        + RegulatoryTableLink + "'>" + RegulatoryTableCaption + "</a></caption><tr><th>Contaminant</th><th>Standard</th></tr>"
        + "<tr><td>Fluoride</td><td>2.0 mg/L</td></tr></table></main>";

    [Fact]
    public void PowerPointKeepsLargeTableCaptionWithFirstPaginatedRow() {
        string rows = string.Concat(Enumerable.Range(1, 12).Select(index =>
            "<tr><td>Row " + index + "</td><td>Standard " + index + "</td></tr>"));
        string html = "<main><h1>Drinking water</h1><table><caption>"
            + RegulatoryTableCaption + "</caption>" + rows + "</table></main>";

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult(
            new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = result.RequireValue();
        PowerPointSlide captionSlide = Assert.Single(presentation.Slides,
            slide => slide.TextBoxes.Any(box => box.Text == RegulatoryTableCaption));
        PowerPointTable firstTable = Assert.Single(captionSlide.Tables);
        Assert.Equal("Row 1", firstTable.GetCell(0, 0).Text);
        Assert.Equal("Standard 1", firstTable.GetCell(0, 1).Text);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.ContentApproximated
            && diagnostic.Detail?.Contains("projection=paginatedNativeTable", StringComparison.Ordinal) == true);
    }

    [Fact]
    public void PowerPointCaptionMetadataLimitDoesNotOmitTheTable() {
        HtmlImportLimits limits = HtmlImportLimits.CreateDefault();
        limits.MaxMetadataCharacters = 12;
        const string html = "<main><h1>Water</h1><table><caption>Caption exceeds metadata limit</caption>"
            + "<tr><th>Term</th><th>Value</th></tr><tr><td>pH</td><td>7</td></tr></table></main>";

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult(
            new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic, Limits = limits });
        using PowerPointPresentation presentation = result.RequireValue();
        Assert.Single(presentation.Slides.SelectMany(slide => slide.Tables));
        Assert.DoesNotContain(presentation.Slides.SelectMany(slide => slide.TextBoxes),
            box => box.Text.Contains("Caption exceeds", StringComparison.Ordinal));
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded
            && diagnostic.Message.Contains("table caption", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void PowerPointAuthoredTableGeometryStaysNativeWithCaption() {
        const string html = "<main><h1>Water</h1><table data-officeimo-top='90' data-officeimo-height='440'>"
            + "<caption>" + RegulatoryTableCaption + "</caption>"
            + "<tr><th>Term</th><th>Value</th></tr><tr><td>pH</td><td>7</td></tr></table></main>";

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult(
            new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = result.RequireValue();
        Assert.Single(presentation.Slides.SelectMany(slide => slide.Tables));
        Assert.Contains(presentation.Slides.SelectMany(slide => slide.TextBoxes),
            box => box.Text == RegulatoryTableCaption);
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic =>
            diagnostic.Message.Contains("split into editable text", StringComparison.OrdinalIgnoreCase));
    }
}
