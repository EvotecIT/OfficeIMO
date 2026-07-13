using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlOfficeAdaptersPowerPointTables {
    [Fact]
    public void PowerPointHtml_RoundTripsMergedTableCells() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointSlide slide = presentation.AddSlide();
        PowerPointTable table = slide.AddTablePoints(3, 3, 70, 90, 360, 150);
        table.GetCell(0, 0).Text = "Merged heading";
        table.GetCell(2, 0).Text = "Tail";
        table.MergeCells(0, 0, 1, 1);
        table.MergeCells(2, 1, 2, 2);

        string html = presentation.ToHtml();
        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult();
        using PowerPointPresentation imported = result.Value;
        PowerPointTable importedTable = Assert.Single(Assert.Single(imported.Slides).Tables);

        Assert.Contains("<td rowspan=\"2\" colspan=\"2\">Merged heading</td>", html, StringComparison.Ordinal);
        Assert.Contains("<td colspan=\"2\"></td>", html, StringComparison.Ordinal);
        Assert.Equal(2, result.MergedRanges);
        Assert.Equal((2, 2), importedTable.GetCell(0, 0).Merge);
        Assert.True(importedTable.GetCell(0, 1).IsMergedCell);
        Assert.Equal((1, 2), importedTable.GetCell(2, 1).Merge);
        Assert.Empty(result.Report.Diagnostics);
    }

    [Fact]
    public void PowerPointHtml_ImportsGenericSpansAndDataAttributeGeometry() {
        const string html = """
            <section class="officeimo-slide">
              <table data-officeimo-left="123" data-officeimo-top="234" data-officeimo-width="345" data-officeimo-height="156">
                <tbody>
                  <tr><th rowspan="2" colspan="2">Group</th><th>Value</th></tr>
                  <tr><td>42</td></tr>
                </tbody>
              </table>
            </section>
            """;

        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult();
        using PowerPointPresentation presentation = result.Value;
        PowerPointTable table = Assert.Single(Assert.Single(presentation.Slides).Tables);

        Assert.Equal(1, result.MergedRanges);
        Assert.Equal((2, 2), table.GetCell(0, 0).Merge);
        Assert.Equal("42", table.GetCell(1, 2).Text);
        Assert.Equal(123D, table.LeftPoints, 3);
        Assert.Equal(234D, table.TopPoints, 3);
        Assert.Equal(345D, table.WidthPoints, 3);
        Assert.Equal(156D, table.HeightPoints, 3);
        Assert.Empty(result.Report.Diagnostics);
    }

    [Fact]
    public void PowerPointHtml_TableCellLimitRejectsOversizedSpanWithoutAllocation() {
        const string html = """
            <section class="officeimo-slide">
              <table><tr><td rowspan="50000" colspan="50000">Value</td></tr></table>
            </section>
            """;

        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult(new HtmlToPowerPointOptions { MaxTableCells = 4 });
        using PowerPointPresentation presentation = result.Value;

        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded);
        Assert.Equal((1, 1), Assert.Single(Assert.Single(presentation.Slides).Tables).GetCell(0, 0).Merge);
    }
}
