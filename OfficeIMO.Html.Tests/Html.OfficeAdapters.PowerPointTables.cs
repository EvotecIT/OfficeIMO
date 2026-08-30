using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

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

        Assert.Contains("<td rowspan=\"2\" colspan=\"2\">", html, StringComparison.Ordinal);
        Assert.Contains("Merged heading", html, StringComparison.Ordinal);
        Assert.Contains("<td colspan=\"2\">", html, StringComparison.Ordinal);
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

    [Fact]
    public void PowerPointHtml_SemanticFormattingUsesTheBoundedNativeTableGrid() {
        const string html = """
            <table>
              <tr><td>First</td><td colspan="999999999999"><strong>Second</strong></td></tr>
            </table>
            """;

        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult(
            new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic, MaxTableCells = 4 });
        using PowerPointPresentation presentation = result.Value;
        PowerPointTable table = Assert.Single(Assert.Single(presentation.Slides).Tables);

        Assert.Equal("First", table.GetCell(0, 0).Text);
        Assert.Equal("Second", table.GetCell(0, 1).Text);
        Assert.True(table.GetCell(0, 1).Runs[0].Bold);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.TableSpanInvalid);
    }

    [Fact]
    public void PowerPointHtml_ExportsApplicableNativeTableStyleTypography() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointSlide slide = presentation.AddSlide();
        const string styleId = "{9A53DA13-207B-4877-931D-000000000240}";
        DocumentFormat.OpenXml.Packaging.PresentationPart presentationPart = slide.SlidePart
            .GetParentParts()
            .OfType<DocumentFormat.OpenXml.Packaging.PresentationPart>()
            .Single();
        PowerPointUtils.CreateTableStylesPart(presentationPart);
        A.TableStyleList styles = presentationPart.TableStylesPart!.TableStyleList!;
        styles.RemoveAllChildren<A.TableStyleEntry>();
        styles.Append(new A.TableStyleEntry(
            $@"<a:tblStyle xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" styleId=""{styleId}"" styleName=""HTML typography"">
  <a:wholeTbl><a:tcTxStyle i=""on""><a:font><a:latin typeface=""Consolas"" /></a:font><a:srgbClr val=""112233"" /></a:tcTxStyle></a:wholeTbl>
  <a:firstRow><a:tcTxStyle b=""on""><a:font><a:latin typeface=""Arial"" /></a:font><a:srgbClr val=""AABBCC"" /></a:tcTxStyle></a:firstRow>
</a:tblStyle>"));

        PowerPointTable table = slide.AddTablePoints(2, 1, 20, 30, 220, 100);
        table.StyleId = styleId;
        table.FirstRow = true;
        table.GetCell(0, 0).Text = "Header";
        table.GetCell(1, 0).Text = "Body";

        string html = presentation.ToHtml();

        Assert.Contains("font-weight:700", html, StringComparison.Ordinal);
        Assert.Contains("font-family:&#39;Arial&#39;", html, StringComparison.Ordinal);
        Assert.Contains("color:#AABBCC", html, StringComparison.Ordinal);
        Assert.Contains("font-style:italic", html, StringComparison.Ordinal);
        Assert.Contains("font-family:&#39;Consolas&#39;", html, StringComparison.Ordinal);
        Assert.Contains("color:#112233", html, StringComparison.Ordinal);
    }
}
