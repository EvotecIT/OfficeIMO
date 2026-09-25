using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests;

public class HtmlOfficeAdaptersPowerPointTables {
    [Fact]
    public void PowerPointHtml_ImportsRoleTableAsEditableGrid() {
        const string html = """
            <main>
              <h1>Water service levels</h1>
              <div role="table" aria-label="Service levels">
                <div role="rowgroup">
                  <div role="row"><div role="columnheader">Term</div><div role="columnheader">Definition</div></div>
                  <div role="row"><div role="cell">Basic water service level</div><div role="cell">Collection time is at most 30 minutes.</div></div>
                </div>
              </div>
            </main>
            """;

        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = result.RequireValue();
        PowerPointTable table = Assert.Single(presentation.Slides.SelectMany(slide => slide.Tables));
        Assert.Equal(presentation.SlideSize.WidthPoints - 128D, table.WidthPoints, 3);
        Assert.Equal("Term", table.GetCell(0, 0).Text);
        Assert.Equal("Definition", table.GetCell(0, 1).Text);
        Assert.Equal("Basic water service level", table.GetCell(1, 0).Text);
        Assert.Equal("Collection time is at most 30 minutes.", table.GetCell(1, 1).Text);
    }

    [Fact]
    public void PowerPointHtml_MovesTallGenericRoleTableToAVisibleSlide() {
        const string html = """
            <main>
              <p>Introductory text before the definitions.</p>
              <div role="table">
                <div role="row"><div role="columnheader">Term</div><div role="columnheader">Definition</div></div>
                <div role="row"><div role="cell">Safely managed</div><div role="cell">Drinking water from an improved source accessible on premises, available when needed, and free from fecal and priority chemical contamination.</div></div>
                <div role="row"><div role="cell">Basic</div><div role="cell">Drinking water from an improved source, provided collection time is not more than 30 minutes round trip, including getting in line and waiting.</div></div>
                <div role="row"><div role="cell">Limited</div><div role="cell">Drinking water from an improved source with collection time exceeding 30 minutes round trip, including getting in line and waiting.</div></div>
                <div role="row"><div role="cell">Unimproved</div><div role="cell">Drinking water from an unprotected dug well or unprotected spring.</div></div>
                <div role="row"><div role="cell">Surface water</div><div role="cell">Drinking water directly from a river, dam, lake, pond, stream, canal, or irrigation canal.</div></div>
              </div>
            </main>
            """;

        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = result.RequireValue();
        PowerPointTable table = Assert.Single(presentation.Slides.SelectMany(slide => slide.Tables));

        Assert.Contains(presentation.Slides.Skip(1), slide => slide.Tables.Contains(table));
        Assert.Equal(30D, table.TopPoints, 3);
        Assert.Equal("Surface water", table.GetCell(5, 0).Text);
        Assert.True(table.TopPoints + table.HeightPoints <= presentation.SlideSize.HeightPoints - 30D);
    }

    [Fact]
    public void PowerPointHtml_SizesUnevenRowsAndAuthoredWidthForGenericTable() {
        string definition = string.Join(" ", Enumerable.Repeat("A longer explanation wraps within a narrow cell.", 4));
        string html = "<div role='table' data-officeimo-width='300'>"
            + "<div role='row'><div role='cell'>Short</div><div role='cell'>Value</div></div>"
            + "<div role='row'><div role='cell'>Long</div><div role='cell'>" + definition + "</div></div></div>";

        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = result.RequireValue();
        PowerPointTable table = Assert.Single(presentation.Slides.SelectMany(slide => slide.Tables));

        Assert.Equal(300D, table.WidthPoints, 3);
        Assert.True(table.GetRowHeightPoints(1) > table.GetRowHeightPoints(0));
        Assert.True(table.TopPoints + table.GetRowHeightPoints(0) + table.GetRowHeightPoints(1)
            <= presentation.SlideSize.HeightPoints - 30D);
        Assert.Equal(definition, table.GetCell(1, 1).Text);
    }

    [Fact]
    public void PowerPointHtml_MeasuresLargeStyledTableRunsBeforePlacingRows() {
        string text = string.Join(" ", Enumerable.Repeat("Large styled cell text wraps across the table.", 3));
        string plainHtml = "<table><tr><td>Label</td><td>" + text + "</td></tr></table>";
        string styledHtml = "<table><tr><td>Label</td><td><span style='font-size:36px'>"
            + text + "</span></td></tr></table>";

        using PowerPointPresentation plain = OfficeIMO.Html.HtmlConversionDocument.Parse(plainHtml)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic }).RequireValue();
        using PowerPointPresentation styled = OfficeIMO.Html.HtmlConversionDocument.Parse(styledHtml)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic }).RequireValue();
        PowerPointTable plainTable = Assert.Single(plain.Slides.SelectMany(slide => slide.Tables));
        PowerPointTable styledTable = Assert.Single(styled.Slides.SelectMany(slide => slide.Tables));

        Assert.True(styledTable.GetRowHeightPoints(0) > plainTable.GetRowHeightPoints(0));
        Assert.True(styledTable.GetCell(0, 1).Runs[0].FontSizePoints >= 27D);
    }

    [Fact]
    public void PowerPointHtml_ReportsAndPaginatesTableTooTallForOneSlide() {
        string rows = string.Concat(Enumerable.Range(0, 15).Select(index =>
            "<div role='row'><div role='cell'>Term " + index + "</div><div role='cell'>"
            + string.Join(" ", Enumerable.Repeat("A long but editable definition.", 4)) + "</div></div>"));
        string html = "<div role='table'>" + rows + "</div>";

        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = result.RequireValue();

        Assert.Empty(presentation.Slides.SelectMany(slide => slide.Tables));
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.ContentApproximated
            && diagnostic.Message.Contains("too tall for one slide", StringComparison.Ordinal));
        Assert.Contains(presentation.Slides.SelectMany(slide => slide.TextBoxes), box =>
            box.Text.Contains("Term 14", StringComparison.Ordinal));
        Assert.All(presentation.Slides.SelectMany(slide => slide.TextBoxes), box =>
            Assert.True(box.TopPoints + box.HeightPoints <= presentation.SlideSize.HeightPoints));
    }

    [Fact]
    public void PowerPointHtml_PreservesRoleTableSpansAndLaterCells() {
        const string html = """
            <div role="table" aria-label="Spanned levels">
              <div role="row"><div role="columnheader" aria-colspan="2">Service</div><div role="columnheader">Definition</div></div>
              <div role="row"><div role="cell" aria-rowspan="2">Basic</div><div role="cell">30 minutes</div><div role="cell">Improved source</div></div>
              <div role="row"><div role="cell">Limited</div><div role="cell">Over 30 minutes</div></div>
            </div>
            """;

        HtmlToPowerPointResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = result.RequireValue();
        PowerPointTable table = Assert.Single(presentation.Slides.SelectMany(slide => slide.Tables));
        Assert.Equal(2, result.MergedRanges);
        Assert.Equal((1, 2), table.GetCell(0, 0).Merge);
        Assert.Equal((2, 1), table.GetCell(1, 0).Merge);
        Assert.Equal("Definition", table.GetCell(0, 2).Text);
        Assert.Equal("Limited", table.GetCell(2, 1).Text);
        Assert.Equal("Over 30 minutes", table.GetCell(2, 2).Text);
    }

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
