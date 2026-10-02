using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Html;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlOneNoteRoleTables {
    [Theory]
    [InlineData("<colgroup><col><col></colgroup>", false)]
    [InlineData("<colgroup span='2'></colgroup>", false)]
    [InlineData("<colgroup><col width='100' style='width:auto'><col></colgroup>", false)]
    [InlineData("<style>col { width:100px }</style><colgroup><col><col></colgroup>", true)]
    [InlineData("<colgroup><col width='100'><col></colgroup>", true)]
    [InlineData("<colgroup><col style='width:100px'><col></colgroup>", true)]
    public void OneNoteHtml_ColumnLossRequiresMeaningfulPresentation(string columns, bool loss) {
        var result = HtmlConversionDocument.Parse("<table>" + columns + "<tr><td>A</td><td>B</td></tr></table>")
            .ToOneNoteSectionResult();
        Assert.Equal(loss, result.Report.HasLoss);
    }

    [Theory]
    [InlineData("width='480'", 10D)]
    [InlineData("width='480' style='width:240px'", 5D)]
    [InlineData("width='480' style='width:auto'", 15D)]
    public void OneNoteHtml_TableWidthHintSurvivesNativeReopen(string attributes, double expected) {
        var result = HtmlConversionDocument.Parse("<table " + attributes + "><tr><td>A</td><td>B</td></tr></table>")
            .ToOneNoteSectionResult();
        var reopened = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(result.RequireValue())));
        var table = Assert.Single(reopened.Pages.SelectMany(p => p.Outlines).SelectMany(o => o.Children).OfType<OneNoteTable>());
        Assert.InRange(table.ColumnWidths.Sum(), expected - 0.01, expected + 0.01);
    }

    [Fact]
    public void OneNoteHtml_ExplicitRowSpanDoesNotOccupyNextRowGroup() {
        var result = HtmlConversionDocument.Parse("<table><tbody><tr><td rowspan='2'>A</td></tr></tbody><tbody><tr><td>B</td></tr></tbody></table>")
            .ToOneNoteSectionResult();
        var reopened = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(result.RequireValue())));
        var table = Assert.Single(reopened.Pages.SelectMany(p => p.Outlines).SelectMany(o => o.Children).OfType<OneNoteTable>());
        Assert.All(table.Rows, row => Assert.Single(row.Cells));
        Assert.Equal("B", string.Concat(table.Rows[1].Cells[0].Content.OfType<OneNoteParagraph>().SelectMany(p => p.Runs).Select(r => r.Text)));
    }

    [Fact]
    public void OneNote_KeepAriaTableRowsEditableAfterSaveAndReload() {
        const string html = "<div role='table' aria-label='Water levels'>"
            + "<div role='row'><div role='columnheader'>Term</div><div role='columnheader'>Definition</div></div>"
            + "<div role='row'><div role='cell'>Basic</div><div role='cell'>At most 30 minutes</div></div>"
            + "</div>";
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);

        HtmlToOneNoteSectionResult oneNoteResult = source.ToOneNoteSectionResult();
        OneNoteSection section = oneNoteResult.RequireValue();
        OneNoteSection reopenedSection = OneNoteSectionReader.Read(
            new MemoryStream(OneNoteSectionWriter.Write(section)));
        OneNoteTable oneNoteTable = Assert.Single(reopenedSection.Pages.SelectMany(page => page.Outlines)
            .SelectMany(outline => outline.Children).OfType<OneNoteTable>());
        Assert.Equal(2, oneNoteTable.Rows.Count);
        Assert.Equal(2, oneNoteTable.Rows[1].Cells.Count);
        Assert.False(oneNoteResult.Report.HasLoss);

    }

    [Fact]
    public void OneNoteHtml_ReportsFlattenedUnevenTableAndSavesRectangularNativeRows() {
        const string html = "<table><tr><th colspan='2'>SI base units</th></tr>"
            + "<tr><td>Length</td><td>meter</td></tr><tr><td>Time</td></tr></table>";

        HtmlToOneNoteSectionResult result = HtmlConversionDocument.Parse(html).ToOneNoteSectionResult();
        OneNoteSection reopened = OneNoteSectionReader.Read(
            new MemoryStream(OneNoteSectionWriter.Write(result.RequireValue())));
        OneNoteTable table = Assert.Single(reopened.Pages.SelectMany(page => page.Outlines)
            .SelectMany(outline => outline.Children).OfType<OneNoteTable>());

        Assert.Equal(3, table.Rows.Count);
        Assert.All(table.Rows, row => Assert.Equal(2, row.Cells.Count));
        Assert.Contains("SI base units", reopened.ToHtmlDocument(), StringComparison.Ordinal);
        Assert.Contains("meter", reopened.ToHtmlDocument(), StringComparison.Ordinal);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.ContentApproximated
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
    }

    [Fact]
    public void OneNoteHtml_FlattenedSpansKeepFollowingCellsInTheirSourceColumns() {
        const string html = "<table><tr><th colspan='2'>Service</th><th>Definition</th></tr>"
            + "<tr><td rowspan='2'>Basic</td><td>30 minutes</td><td>Improved source</td></tr>"
            + "<tr><td>Limited</td><td>Over 30 minutes</td></tr></table>";

        HtmlToOneNoteSectionResult result = HtmlConversionDocument.Parse(html).ToOneNoteSectionResult();
        OneNoteSection reopened = OneNoteSectionReader.Read(
            new MemoryStream(OneNoteSectionWriter.Write(result.RequireValue())));
        OneNoteTable table = Assert.Single(reopened.Pages.SelectMany(page => page.Outlines)
            .SelectMany(outline => outline.Children).OfType<OneNoteTable>());

        Assert.All(table.Rows, row => Assert.Equal(3, row.Cells.Count));
        Assert.Equal("Service", CellText(table.Rows[0].Cells[0]));
        Assert.Equal(string.Empty, CellText(table.Rows[0].Cells[1]));
        Assert.Equal("Definition", CellText(table.Rows[0].Cells[2]));
        Assert.Equal("Basic", CellText(table.Rows[1].Cells[0]));
        Assert.Equal("30 minutes", CellText(table.Rows[1].Cells[1]));
        Assert.Equal("Improved source", CellText(table.Rows[1].Cells[2]));
        Assert.Equal(string.Empty, CellText(table.Rows[2].Cells[0]));
        Assert.Equal("Limited", CellText(table.Rows[2].Cells[1]));
        Assert.Equal("Over 30 minutes", CellText(table.Rows[2].Cells[2]));
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.ContentApproximated
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation);

        static string CellText(OneNoteTableCell cell) => string.Concat(cell.Content
            .OfType<OneNoteParagraph>().SelectMany(paragraph => paragraph.Runs).Select(run => run.Text));
    }

    [Fact]
    public void OneNoteHtml_OversizedTableLeavesBudgetForFollowingValidTable() {
        const string html = "<table><tr><td>Too</td><td>many</td></tr><tr><td>cells</td><td>here</td></tr></table>"
            + "<table><tr><td>Retained</td></tr></table>";
        HtmlImportLimits limits = HtmlImportLimits.CreateDefault();
        limits.MaxTables = 1;
        limits.MaxTableCells = 3;

        HtmlToOneNoteSectionResult result = HtmlConversionDocument.Parse(html).ToOneNoteSectionResult(
            new HtmlToOneNoteOptions { Limits = limits });
        OneNoteSection reopened = OneNoteSectionReader.Read(
            new MemoryStream(OneNoteSectionWriter.Write(result.RequireValue())));
        OneNoteTable table = Assert.Single(reopened.Pages.SelectMany(page => page.Outlines)
            .SelectMany(outline => outline.Children).OfType<OneNoteTable>());

        Assert.Equal(1, result.Tables);
        Assert.Equal("Retained", string.Concat(table.Rows[0].Cells[0].Content
            .OfType<OneNoteParagraph>().SelectMany(paragraph => paragraph.Runs).Select(run => run.Text)));
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
    }

    [Fact]
    public void OneNoteHtml_RowSpanAcrossEmptySourceRowDoesNotShiftLaterCells() {
        const string html = "<table><tr><td rowspan='2'>A</td><td>B</td></tr><tr></tr>"
            + "<tr><td>C</td><td>D</td></tr></table>";
        HtmlSemanticTable semantic = Assert.Single(HtmlConversionDocument.Parse(html)
            .CreateSemanticDocumentForConversion(HtmlCssMediaContext.Screen).RootTables).Table!;
        Assert.Equal(new[] { 0, 2 }, semantic.Rows.Select(row => row.SourceRowIndex));

        HtmlToOneNoteSectionResult result = HtmlConversionDocument.Parse(html).ToOneNoteSectionResult();
        OneNoteSection reopened = OneNoteSectionReader.Read(
            new MemoryStream(OneNoteSectionWriter.Write(result.RequireValue())));
        OneNoteTable table = Assert.Single(reopened.Pages.SelectMany(page => page.Outlines)
            .SelectMany(outline => outline.Children).OfType<OneNoteTable>());

        Assert.Equal(2, table.Rows.Count);
        Assert.All(table.Rows, row => Assert.Equal(2, row.Cells.Count));
        Assert.Equal("C", string.Concat(table.Rows[1].Cells[0].Content
            .OfType<OneNoteParagraph>().SelectMany(paragraph => paragraph.Runs).Select(run => run.Text)));
        Assert.Equal("D", string.Concat(table.Rows[1].Cells[1].Content
            .OfType<OneNoteParagraph>().SelectMany(paragraph => paragraph.Runs).Select(run => run.Text)));
    }

    [Fact]
    public void OneNoteHtml_AuthoredColumnWidthsSurviveNativeReopen() {
        var result = HtmlConversionDocument.Parse("<table style='width:360px'><tr><td style='width:120px'>A</td><td style='width:240px'>B</td></tr></table>")
            .ToOneNoteSectionResult();
        var reopened = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(result.RequireValue())));
        var table = Assert.Single(reopened.Pages.SelectMany(p => p.Outlines).SelectMany(o => o.Children).OfType<OneNoteTable>());
        Assert.Equal(2, table.ColumnWidths.Count);
        Assert.InRange(table.ColumnWidths[0], 2.49, 2.51);
        Assert.InRange(table.ColumnWidths[1], 4.99, 5.01);
    }

    [Fact]
    public void OneNoteHtml_LinkedCaptionKeepsTextWhenLinkExceedsMetadataLimit() {
        var limits = HtmlImportLimits.CreateDefault(); limits.MaxMetadataCharacters = 256;
        var result = HtmlConversionDocument.Parse("<table><caption><a href='https://example.org/" + new string('x', 300)
            + "'>Retained caption</a></caption><tr><td>Cell</td></tr></table>")
            .ToOneNoteSectionResult(new HtmlToOneNoteOptions { Limits = limits });
        var reopened = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(result.RequireValue())));
        var caption = Assert.Single(reopened.Pages.SelectMany(p => p.Outlines).SelectMany(o => o.Children).OfType<OneNoteParagraph>());
        Assert.Equal("Retained caption", string.Concat(caption.Runs.Select(r => r.Text)));
        Assert.All(caption.Runs, r => Assert.True(string.IsNullOrEmpty(r.Hyperlink)));
        Assert.Contains(result.Report.Diagnostics, d => d.Code == HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded);
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OneNoteHtml_ImageLinkSurvivesNativeReopenWithinMetadataLimit(bool oversized) {
        const string png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aX1sAAAAASUVORK5CYII=";
        string link = "https://example.org/" + (oversized ? new string('x', 300) : "source");
        var limits = HtmlImportLimits.CreateDefault(); limits.MaxMetadataCharacters = 256;
        var result = HtmlConversionDocument.Parse("<a href='" + link + "'><img src='data:image/png;base64," + png + "' alt='Photo'></a>")
            .ToOneNoteSectionResult(new HtmlToOneNoteOptions { Limits = limits });
        var reopened = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(result.RequireValue())));
        var image = Assert.Single(reopened.Pages.SelectMany(p => p.Outlines).SelectMany(o => o.Children).OfType<OneNoteImage>());
        Assert.Equal("Photo", image.AltText);
        if (oversized) Assert.True(string.IsNullOrEmpty(image.Hyperlink));
        else Assert.Equal(link, image.Hyperlink);
    }
}
