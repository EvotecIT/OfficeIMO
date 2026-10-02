using OfficeIMO.Html;
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Html;
using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlOfficeAdapters {
    [Fact]
    public void OneNoteAndRtf_KeepAriaTableRowsEditableAfterSaveAndReload() {
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

        HtmlToRtfResult rtfResult = source.ToRtfDocumentResult();
        RtfDocument reloaded = RtfDocument.Load(rtfResult.RequireValue().ToBytes());
        RtfTable rtfTable = Assert.Single(reloaded.Blocks.OfType<RtfTable>());
        Assert.Equal(2, rtfTable.Rows.Count);
        Assert.Equal(2, rtfTable.Rows[1].Cells.Count);
        Assert.Contains("At most 30 minutes", reloaded.ToHtml(), StringComparison.Ordinal);
        Assert.False(rtfResult.Report.HasLoss);
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
    public void Rtf_ReportsUnsupportedAriaTableStructureAsApproximation() {
        const string html = "<div role='table'><p>Unstructured</p><div role='row'><div role='cell'>Value</div></div></div>";

        HtmlToRtfResult result = HtmlConversionDocument.Parse(html).ToRtfDocumentResult();

        Assert.Empty(result.RequireValue().Blocks.OfType<RtfTable>());
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.ContentApproximated
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
    }

    [Fact]
    public void Rtf_ReportsActiveStylesheetsItDoesNotApply() {
        const string html = "<html><head><link rel='stylesheet' href='cid:layout.css'>"
            + "<link rel='stylesheet' disabled href='cid:disabled.css'>"
            + "<link rel='alternate stylesheet' title='alternate' href='cid:alternate.css'>"
            + "<link rel='stylesheet' type='text/less' href='cid:not-css.css'>"
            + "<style>p { color: red; }</style>"
            + "<style type='text/less'>p { color: blue; }</style>"
            + "</head><body><p>Content</p></body></html>";

        HtmlToRtfResult result = HtmlConversionDocument.Parse(html).ToRtfDocumentResult();

        Assert.Equal("cid:layout.css", Assert.Single(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == "HtmlStylesheetLinkSkipped").Source);
        Assert.Equal(OfficeConversionLossKind.Omission, Assert.Single(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == "HtmlStylesheetElementSkipped").LossKind);
    }

    [Fact]
    public void Rtf_AriaRoleTable_SyntheticNodesDoNotConsumeSourceLimits() {
        const string html = "<div role='table'><div role='row'><div role='cell'></div></div></div>";

        HtmlToRtfResult result = HtmlConversionDocument.Parse(html).ToRtfDocumentResult(
            new HtmlToRtfOptions { MaxHtmlNodes = 6, MaxHtmlDepth = 5 });

        Assert.Single(result.RequireValue().Blocks.OfType<RtfTable>());
        HtmlRtfConversionLimitException exception = Assert.Throws<HtmlRtfConversionLimitException>(() =>
            HtmlConversionDocument.Parse(html).ToRtfDocumentResult(new HtmlToRtfOptions { MaxHtmlDepth = 4 }));
        Assert.Equal(nameof(HtmlToRtfOptions.MaxHtmlDepth), exception.LimitSource);
    }

}
