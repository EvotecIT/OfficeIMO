using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
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

    [Fact]
    public void ExcelHtml_ImportsAriaTableAsPairedEditableCells() {
        const string html = """
            <main>
              <h1>Water service levels</h1>
              <div role="table" aria-label="Service levels">
                <div role="rowgroup">
                  <div role="row"><div role="columnheader">Term</div><div role="columnheader">Definition</div></div>
                  <div role="row"><div role="cell"><p>Basic water service level</p></div><div role="cell"><p>Collection time is at most 30 minutes.</p></div></div>
                  <div role="row"><div role="cell">Limited water service level</div><div role="cell">Collection time exceeds 30 minutes.</div></div>
                </div>
              </div>
            </main>
            """;

        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        HtmlSemanticDocument semantic = document.CreateSemanticDocumentForConversion(HtmlCssMediaContext.Screen);
        HtmlSemanticTable table = Assert.Single(semantic.RootTables).Table!;
        Assert.Equal(3, table.Rows.Count);
        Assert.Equal(2, table.Rows[1].Cells.Count);
        Assert.Equal("Basic water service level", table.Rows[1].Cells[0].Text);
        Assert.Equal("Collection time is at most 30 minutes.", table.Rows[1].Cells[1].Text);

        HtmlToExcelResult result = document.ToExcelDocumentResult(new HtmlToExcelOptions {
            Mode = HtmlImportMode.Generic
        });
        using ExcelDocument workbook = result.RequireValue();
        ExcelSheet sheet = Assert.Single(workbook.Sheets, candidate => candidate.Name == "Service levels");
        Assert.True(sheet.TryGetCellValueSnapshot(2, 1, out ExcelCellValueSnapshot? term));
        Assert.True(sheet.TryGetCellValueSnapshot(2, 2, out ExcelCellValueSnapshot? definition));
        Assert.Equal("Basic water service level", term!.Text);
        Assert.Equal("Collection time is at most 30 minutes.", definition!.Text);
    }

    [Fact]
    public void ExcelHtml_GenericTermDefinitionTableSavesReadableColumnsAndRows() {
        const string html = "<div role='table' aria-label='Water levels'>"
            + "<div role='row'><div role='columnheader'>Term</div><div role='columnheader'>Definition</div></div>"
            + "<div role='row'><div role='cell'>Safely managed water service level</div>"
            + "<div role='cell'>Drinking water from an improved source accessible on premises, available when needed, and free from fecal and priority chemical contamination.</div></div>"
            + "</div>";

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(new HtmlToExcelOptions {
            Mode = HtmlImportMode.Generic
        });
        using ExcelDocument workbook = result.RequireValue();
        using MemoryStream artifact = workbook.ToStream();
        byte[] bytes = artifact.ToArray();
        using (SpreadsheetDocument package = SpreadsheetDocument.Open(new MemoryStream(bytes), false)) {
            WorksheetPart worksheetPart = Assert.Single(package.WorkbookPart!.WorksheetParts);
            Column[] columns = worksheetPart.Worksheet.GetFirstChild<Columns>()!.Elements<Column>().ToArray();
            Assert.Equal(2, columns.Length);
            Assert.InRange(columns[0].Width!.Value, 20D, 40D);
            Assert.InRange(columns[1].Width!.Value, 35D, 55D);
            Assert.True(columns.Sum(column => column.Width!.Value) <= 80D);
            Row[] rows = worksheetPart.Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().ToArray();
            Cell definition = Assert.Single(rows[1].Elements<Cell>(), cell => cell.CellReference!.Value == "B2");
            CellFormat format = package.WorkbookPart.WorkbookStylesPart!.Stylesheet.CellFormats!
                .Elements<CellFormat>().ElementAt((int)definition.StyleIndex!.Value);
            Assert.True(format.Alignment?.WrapText?.Value);
        }

        using ExcelDocument reopened = ExcelDocument.Load(new MemoryStream(bytes));
        ExcelSheet sheet = Assert.Single(reopened.Sheets);
        Assert.Equal("Safely managed water service level", sheet.CellAt(2, 1).GetValue<string>());
        Assert.Contains("priority chemical contamination", sheet.CellAt(2, 2).GetValue<string>(), StringComparison.Ordinal);
    }

    [Fact]
    public void ExcelHtml_GenericTablePresentationFollowsImportedRowsAcrossAnEmptySourceRow() {
        const string html = "<table><tr><th>Term</th><th>Definition</th></tr><tr></tr>"
            + "<tr><td>Safely managed water service level</td>"
            + "<td>Drinking water from an improved source accessible on premises, available when needed.</td></tr></table>";

        using ExcelDocument workbook = HtmlConversionDocument.Parse(html).ToExcelDocument(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using MemoryStream artifact = workbook.ToStream();
        using SpreadsheetDocument package = SpreadsheetDocument.Open(artifact, false);
        WorksheetPart worksheet = Assert.Single(package.WorkbookPart!.WorksheetParts);
        Cell[] cells = worksheet.Worksheet.Descendants<Cell>().ToArray();
        Assert.DoesNotContain(cells, cell => cell.CellReference!.Value is "A2" or "B2");
        Cell definition = Assert.Single(cells, cell => cell.CellReference!.Value == "B3");
        CellFormat format = package.WorkbookPart.WorkbookStylesPart!.Stylesheet.CellFormats!
            .Elements<CellFormat>().ElementAt((int)definition.StyleIndex!.Value);
        Assert.True(format.Alignment?.WrapText?.Value);
    }

    [Fact]
    public void ExcelHtml_GenericTablePresentationDoesNotCreateCellsBeyondTheImportLimit() {
        const string html = "<table><tr><th>Term</th><th>Definition</th></tr>"
            + "<tr><td>Safely managed water service level</td><td>Omitted definition</td></tr></table>";

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Generic, MaxTableCells = 3 });
        using ExcelDocument workbook = result.RequireValue();
        using MemoryStream artifact = workbook.ToStream();
        using SpreadsheetDocument package = SpreadsheetDocument.Open(artifact, false);
        WorksheetPart worksheet = Assert.Single(package.WorkbookPart!.WorksheetParts);
        string[] references = worksheet.Worksheet.Descendants<Cell>()
            .Select(cell => cell.CellReference!.Value!).ToArray();
        Assert.Equal(new[] { "A1", "B1", "A2" }, references);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded);
    }

    [Fact]
    public void ExcelHtml_PreservesAriaTableSpansWithoutShiftingCells() {
        const string html = """
            <div role="table" aria-label="Spanned levels">
              <div role="row"><div role="columnheader" aria-colspan="2">Service</div><div role="columnheader">Definition</div></div>
              <div role="row"><div role="cell" aria-rowspan="2">Basic</div><div role="cell">30 minutes</div><div role="cell">Improved source</div></div>
              <div role="row"><div role="cell">Limited</div><div role="cell">Over 30 minutes</div></div>
            </div>
            """;

        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        HtmlSemanticTable table = Assert.Single(document
            .CreateSemanticDocumentForConversion(HtmlCssMediaContext.Screen).RootTables).Table!;
        Assert.Equal(2, table.Rows[0].Cells[0].ColumnSpan);
        Assert.Equal(2, table.Rows[1].Cells[0].RowSpan);

        HtmlToExcelResult result = document.ToExcelDocumentResult(new HtmlToExcelOptions {
            Mode = HtmlImportMode.Generic
        });
        using ExcelDocument workbook = result.RequireValue();
        ExcelSheet sheet = Assert.Single(workbook.Sheets, candidate => candidate.Name == "Spanned levels");
        Assert.Equal(2, result.MergedRanges);
        Assert.True(sheet.TryGetCellValueSnapshot(1, 3, out ExcelCellValueSnapshot? heading));
        Assert.Equal("Definition", heading!.Text);
        Assert.True(sheet.TryGetCellValueSnapshot(3, 2, out ExcelCellValueSnapshot? term));
        Assert.True(sheet.TryGetCellValueSnapshot(3, 3, out ExcelCellValueSnapshot? definition));
        Assert.Equal("Limited", term!.Text);
        Assert.Equal("Over 30 minutes", definition!.Text);
    }
}
