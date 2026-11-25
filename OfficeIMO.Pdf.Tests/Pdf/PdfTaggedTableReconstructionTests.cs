using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Excel;
using OfficeIMO.Pdf.Word;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTaggedTableReconstructionTests {
    private const double Left = 60D;
    private const double ColumnWidth = 100D;

    [Fact]
    public void TaggedCellsKeepTheirGridThroughSpansWrappedLinesAndEmptyCells() {
        byte[] bytes = CreateTaggedDocument(canvas => canvas.Structure(PdfCanvasStructureRole.Table, table => table
            .Structure(PdfCanvasStructureRole.TableBody, head => head
                .Structure(PdfCanvasStructureRole.TableRow, row => row
                    .Structure(PdfCanvasStructureRole.TableHeaderCell, cell => Lines(cell, 0, 100D, "Region"))
                    .Structure(PdfCanvasStructureRole.TableHeaderCell, cell => Lines(cell, 1, 100D, "Q1"))
                    .Structure(PdfCanvasStructureRole.TableHeaderCell, cell => Lines(cell, 2, 100D, "Q2"))
                    .Structure(PdfCanvasStructureRole.TableHeaderCell, cell => Lines(cell, 3, 100D, "Notes"))),
                new PdfCanvasStructureOptions { TableRowGroupType = "THead" })
            .Structure(PdfCanvasStructureRole.TableBody, body => body
                .Structure(PdfCanvasStructureRole.TableRow, row => row
                    .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, 0, 120D, "North"))
                    .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, 1, 120D, "Closed for the quarter"),
                        new PdfCanvasStructureOptions { ColumnSpan = 2 })
                    .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, 3, 120D, "Line one", "line two")))
                .Structure(PdfCanvasStructureRole.TableRow, row => row
                    .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, 0, 150D, "South"),
                        new PdfCanvasStructureOptions { RowSpan = 2 })
                    .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, 1, 150D, "10"))
                    .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, 2, 150D, string.Empty))
                    .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, 3, 150D, "Short")))
                .Structure(PdfCanvasStructureRole.TableRow, row => row
                    .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, 1, 170D, "12"))
                    .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, 2, 170D, "14"))
                    .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, 3, 170D, "Final"))))));

        using PdfDocument document = PdfDocument.Load(new MemoryStream(bytes));
        PdfDocumentReadResult read = document.Read();

        PdfLogicalTable table = Assert.Single(read.Pages[0].Tables);
        Assert.Equal("tagged-structure", table.DetectionKind);
        Assert.Collection(table.Rows,
            row => Assert.Equal(new[] { "Region", "Q1", "Q2", "Notes" }, row),
            row => Assert.Equal(new[] { "North", "Closed for the quarter", "", "Line one line two" }, row),
            row => Assert.Equal(new[] { "South", "10", "", "Short" }, row),
            row => Assert.Equal(new[] { "", "12", "14", "Final" }, row));
        PdfLogicalTableCell merged = table.Cells.Single(cell => cell.RowIndex == 1 && cell.ColumnIndex == 1);
        Assert.Equal(2, merged.ColumnSpan);
        Assert.True(table.Cells.Single(cell => cell.RowIndex == 1 && cell.ColumnIndex == 2).IsCoveredBySpan);
        Assert.Equal(2, table.Cells.Single(cell => cell.RowIndex == 2 && cell.ColumnIndex == 0).RowSpan);
        Assert.True(table.Cells.Single(cell => cell.RowIndex == 3 && cell.ColumnIndex == 0).IsCoveredBySpan);
        Assert.False(table.Cells.Single(cell => cell.RowIndex == 2 && cell.ColumnIndex == 2).IsCoveredBySpan);
        Assert.Equal(new[] { "Region", "Q1", "Q2", "Notes" }, PdfLogicalTableAnalysis.Extract(table).Columns);
        Assert.Empty(read.TaggedTableDiagnostics);

        Assert.Contains("| Region | Q1 | Q2 | Notes |", read.ExtractMarkdown(), StringComparison.Ordinal);
        Assert.Contains("| North | Closed for the quarter |  | Line one line two |", read.ExtractMarkdown(), StringComparison.Ordinal);

        using OfficeIMO.Word.WordDocument word = read.ConvertToWordWithEvidence().Value;
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(word.ToBytes()), false);
        Table wordTable = Assert.Single(package.MainDocumentPart!.Document!.Body!.Elements<Table>());
        TableRow[] wordRows = wordTable.Elements<TableRow>().ToArray();
        Assert.Equal(4, wordRows.Length);
        TableCell spanned = wordRows[1].Elements<TableCell>().ElementAt(1);
        Assert.Equal(2, spanned.TableCellProperties!.GridSpan!.Val!.Value);
        Assert.Equal("Closed for the quarter", spanned.InnerText);
        Assert.Equal(MergedCellValues.Restart, wordRows[2].Elements<TableCell>().First().TableCellProperties!.VerticalMerge!.Val!.Value);
        Assert.NotNull(wordRows[3].Elements<TableCell>().First().TableCellProperties!.VerticalMerge);
    }

    [Fact]
    public void TaggedTableOnTwoPagesStaysOneLogicalTable() {
        // Rows sit mid-page, away from the page edges that geometric continuation needs,
        // so only the shared Table element can join the two pages.
        byte[] bytes = CreateTaggedDocument(
            canvas => TablePage(canvas, ("Alpha", "1"), ("Beta", "2"), ("Gamma", "3")),
            canvas => TablePage(canvas, ("Delta", "4"), ("Epsilon", "5")));

        using PdfDocument document = PdfDocument.Load(new MemoryStream(bytes));
        PdfDocumentReadResult read = document.Read();

        Assert.Equal(2, read.Pages.Count);
        Assert.Equal("tagged-structure", Assert.Single(read.Pages[0].Tables).DetectionKind);
        Assert.Equal("tagged-structure", Assert.Single(read.Pages[1].Tables).DetectionKind);
        PdfLogicalTableContinuationGroup group = Assert.Single(read.GetTableContinuationGroups());
        Assert.Equal(2, group.Segments.Count);
        Assert.True(group.Evidence.HasFlag(PdfLogicalTableContinuationEvidence.TaggedStructure));
        Assert.Equal(new[] { "Item", "Count" }, group.Columns);
        Assert.Equal(
            new[] { "Alpha|1", "Beta|2", "Gamma|3", "Delta|4", "Epsilon|5" },
            group.Rows.Select(static row => string.Join("|", row)));

        PdfExcelTableImportResult excel = read.ConvertToExcelWithEvidence();
        using (excel.Value) {
            PdfExcelTableImportEntry entry = Assert.Single(excel.Report.Entries);
            Assert.Equal(new[] { 1, 2 }, entry.SourcePageNumbers);
            Assert.Equal(5, entry.RowCount);
        }
    }

    [Fact]
    public void BrokenTableTagsFallBackToGeometryWithDiagnostic() {
        byte[] bytes = CreateTaggedDocument(canvas => TablePage(canvas, ("Alpha", "1"), ("Beta", "2"), ("Gamma", "3")));
        using (PdfDocument intact = PdfDocument.Load(new MemoryStream(bytes))) {
            Assert.Equal("tagged-structure", Assert.Single(intact.Read().Pages[0].Tables).DetectionKind);
        }

        // Renumber the last marked-content sequence in the page content, so its cell's
        // structure reference points at an MCID the page no longer has.
        byte[] broken = RenumberLastContentMcid(bytes);
        using PdfDocument document = PdfDocument.Load(new MemoryStream(broken));
        PdfDocumentReadResult read = document.Read();

        PdfTaggedTableDiagnostic diagnostic = Assert.Single(read.TaggedTableDiagnostics);
        Assert.Equal(PdfTaggedTableFallbackReason.UnresolvedContent, diagnostic.Reason);
        Assert.Equal(1, diagnostic.PageNumber);
        Assert.DoesNotContain(read.Pages[0].Tables, static table => table.DetectionKind == "tagged-structure");
        foreach (string text in new[] { "Item", "Count", "Alpha", "Beta", "Gamma", "3" }) {
            Assert.Contains(text, read.Text, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void CellsTaggedAgainstTheirPageOrderFallBackToGeometry() {
        // The second body row tags Count before Item while the page shows Item on the left.
        byte[] bytes = CreateTaggedDocument(canvas => TablePage(canvas, new[] { ("Alpha", "1"), ("Beta", "2"), ("Gamma", "3") }, swappedRow: 1));
        using PdfDocument document = PdfDocument.Load(new MemoryStream(bytes));
        PdfDocumentReadResult read = document.Read();

        PdfTaggedTableDiagnostic diagnostic = Assert.Single(read.TaggedTableDiagnostics);
        Assert.Equal(PdfTaggedTableFallbackReason.CellOrderConflictsWithLayout, diagnostic.Reason);
        Assert.DoesNotContain(read.Pages[0].Tables, static table => table.DetectionKind == "tagged-structure");
        Assert.DoesNotContain(read.Pages[0].Tables.SelectMany(static table => table.Rows),
            static row => row.Count == 2 && row[0] == "2" && row[1] == "Beta");
    }

    [Fact]
    public void PageSelectionWithoutTheTablePageIgnoresTheTable() {
        byte[] bytes = CreateTaggedDocument(
            canvas => canvas.Structure(PdfCanvasStructureRole.Paragraph, paragraph => Lines(paragraph, 0, 100D, "Introduction")),
            canvas => TablePage(canvas, ("Alpha", "1"), ("Beta", "2")));
        using PdfDocument document = PdfDocument.Load(new MemoryStream(bytes));

        PdfDocumentReadResult read = document.Read(new PdfLoadOptions { PageSelection = PdfPageSelection.From(1) });

        PdfLogicalPage page = Assert.Single(read.Pages);
        Assert.Empty(page.Tables);
        Assert.Empty(read.TaggedTableDiagnostics);
        Assert.Contains("Introduction", read.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void StructuralFailureInTableAcrossPagesIsReportedOnEachPage() {
        // A paragraph directly under Table is not a row group, row or cell. The Table element
        // spans both pages, so it has no /Pg of its own.
        byte[] bytes = CreateTaggedDocument(
            canvas => TablePage(canvas, ("Alpha", "1"), ("Beta", "2")),
            canvas => TablePage(canvas, new[] { ("Gamma", "3"), ("Delta", "4") }, strayParagraph: "Stray"));
        using PdfDocument document = PdfDocument.Load(new MemoryStream(bytes));
        PdfDocumentReadResult read = document.Read();

        Assert.Equal(new[] { 1, 2 }, read.TaggedTableDiagnostics.Select(static item => item.PageNumber));
        Assert.All(read.TaggedTableDiagnostics,
            static item => Assert.Equal(PdfTaggedTableFallbackReason.UnsupportedStructure, item.Reason));
        Assert.DoesNotContain(read.Pages.SelectMany(static page => page.Tables), static table => table.DetectionKind == "tagged-structure");
    }

    [Fact]
    public void WordExportSizesEveryColumnWhenTheHeaderRowSpansCells() {
        byte[] bytes = CreateTaggedDocument(canvas => canvas.Structure(PdfCanvasStructureRole.Table, table => table
            .Structure(PdfCanvasStructureRole.TableBody, head => head
                .Structure(PdfCanvasStructureRole.TableRow, row => row
                    .Structure(PdfCanvasStructureRole.TableHeaderCell, cell => Lines(cell, 0, 100D, "Region"))
                    .Structure(PdfCanvasStructureRole.TableHeaderCell, cell => Lines(cell, 1, 100D, "First half"),
                        new PdfCanvasStructureOptions { ColumnSpan = 2 })),
                new PdfCanvasStructureOptions { TableRowGroupType = "THead" })
            .Structure(PdfCanvasStructureRole.TableBody, body => body
                .Structure(PdfCanvasStructureRole.TableRow, row => row
                    .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, 0, 120D, "North"))
                    .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, 1, 120D, "10"))
                    .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, 2, 120D, "12"))))));
        using PdfDocument document = PdfDocument.Load(new MemoryStream(bytes));
        PdfDocumentReadResult read = document.Read();

        using OfficeIMO.Word.WordDocument word = read.ConvertToWordWithEvidence().Value;
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(word.ToBytes()), false);
        Table wordTable = Assert.Single(package.MainDocumentPart!.Document!.Body!.Elements<Table>());
        TableRow[] wordRows = wordTable.Elements<TableRow>().ToArray();
        TableCell spanned = wordRows[0].Elements<TableCell>().ElementAt(1);
        Assert.Equal(2, spanned.TableCellProperties!.GridSpan!.Val!.Value);
        Assert.Equal("First half", spanned.InnerText);
        Assert.Equal(new[] { "North", "10", "12" }, wordRows[1].Elements<TableCell>().Select(static cell => cell.InnerText));
        Assert.All(wordRows[1].Elements<TableCell>(),
            static cell => Assert.False(string.IsNullOrEmpty(cell.TableCellProperties?.TableCellWidth?.Width?.Value)));
    }

    [Fact]
    public void IdenticalFirstRowsOfTaggedPagesStayData() {
        byte[] bytes = CreateTaggedDocument(
            canvas => TablePage(canvas, ("Alpha", "Active"), ("Beta", "Paused")),
            canvas => TablePage(canvas, ("Alpha", "Active"), ("Gamma", "Paused")));
        using PdfDocument document = PdfDocument.Load(new MemoryStream(bytes));

        PdfLogicalTableContinuationGroup group = Assert.Single(document.Read().GetTableContinuationGroups());

        Assert.Equal(new[] { "Item", "Count" }, group.Columns);
        Assert.Equal(
            new[] { "Alpha|Active", "Beta|Paused", "Alpha|Active", "Gamma|Paused" },
            group.Rows.Select(static row => string.Join("|", row)));
    }

    [Fact]
    public void SelectedContinuationPageKeepsTheTaggedHeaderNames() {
        // The header is tagged only on page 1; page 2 continues the same Table element.
        byte[] bytes = CreateTaggedDocument(
            canvas => TablePage(canvas, ("Alpha", "1"), ("Beta", "2")),
            canvas => TablePage(canvas, new[] { ("Gamma", "3"), ("Delta", "4") }, header: false));
        using PdfDocument document = PdfDocument.Load(new MemoryStream(bytes));

        PdfDocumentReadResult read = document.Read(new PdfLoadOptions { PageSelection = PdfPageSelection.From(2) });

        PdfLogicalTable table = Assert.Single(Assert.Single(read.Pages).Tables);
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);
        Assert.Equal(new[] { "Item", "Count" }, data.Columns);
        Assert.Equal(new[] { "Gamma|3", "Delta|4" }, data.Rows.Select(static row => string.Join("|", row)));
    }

    [Fact]
    public void StructureTreeWithoutMarkedCatalogKeepsGeometricTables() {
        string tagged = PdfEncoding.Latin1GetString(CreateTaggedDocument(canvas => TablePage(canvas, ("Alpha", "1"), ("Beta", "2"), ("Gamma", "3"))));
        Assert.Contains("/Marked true", tagged, StringComparison.Ordinal);
        // Same length keeps every cross-reference offset valid; /MarkInfo no longer declares the file tagged.
        byte[] bytes = PdfEncoding.Latin1GetBytes(tagged.Replace("/Marked true", "/Marker true"));

        using PdfDocument document = PdfDocument.Load(new MemoryStream(bytes));
        PdfDocumentReadResult read = document.Read();

        Assert.DoesNotContain(read.Pages[0].Tables, static table => table.DetectionKind == "tagged-structure");
        Assert.Empty(read.TaggedTableDiagnostics);
    }

    private static void TablePage(PdfPageCanvas canvas, params (string Item, string Count)[] rows) =>
        TablePage(canvas, rows, swappedRow: -1);

    private static void TablePage(PdfPageCanvas canvas, (string Item, string Count)[] rows, int swappedRow = -1,
        string? strayParagraph = null, bool header = true) {
        var tableKey = new PdfCanvasStructureOptions { StructureElementKey = "table" };
        var headKey = new PdfCanvasStructureOptions { StructureElementKey = "head", TableRowGroupType = "THead" };
        var headRowKey = new PdfCanvasStructureOptions { StructureElementKey = "head-row" };
        var bodyKey = new PdfCanvasStructureOptions { StructureElementKey = "body" };
        canvas.Structure(PdfCanvasStructureRole.Table, table => {
            if (header) table.Structure(PdfCanvasStructureRole.TableBody, head => head
                .Structure(PdfCanvasStructureRole.TableRow, row => row
                    .Structure(PdfCanvasStructureRole.TableHeaderCell, cell => Lines(cell, 0, 300D, "Item"),
                        new PdfCanvasStructureOptions { StructureElementKey = "item" })
                    .Structure(PdfCanvasStructureRole.TableHeaderCell, cell => Lines(cell, 1, 300D, "Count"),
                        new PdfCanvasStructureOptions { StructureElementKey = "count" }), headRowKey), headKey);
            table.Structure(PdfCanvasStructureRole.TableBody, body => {
                for (int index = 0; index < rows.Length; index++) {
                    (string item, string count) = rows[index];
                    double y = 320D + (index * 20D);
                    bool swapped = index == swappedRow;
                    body.Structure(PdfCanvasStructureRole.TableRow, row => row
                        .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, swapped ? 1 : 0, y, swapped ? count : item))
                        .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, swapped ? 0 : 1, y, swapped ? item : count)));
                }
            }, bodyKey);
            if (strayParagraph is not null) {
                table.Structure(PdfCanvasStructureRole.Paragraph, paragraph => Lines(paragraph, 0, 500D, strayParagraph));
            }
        }, tableKey);
    }

    private static void Lines(PdfPageCanvas cell, int column, double top, params string[] lines) {
        for (int index = 0; index < lines.Length; index++) {
            cell.Text(lines[index],
                OfficeIMO.OfficeLength.Points(Left + (column * ColumnWidth)),
                OfficeIMO.OfficeLength.Points(top + (index * 11D)),
                OfficeIMO.OfficeLength.Points(ColumnWidth * 2D),
                OfficeIMO.OfficeLength.Points(11D));
        }
    }

    private static byte[] CreateTaggedDocument(params Action<PdfPageCanvas>[] pages) {
        var builder = PdfDocument.Create(new PdfCreateOptions { CompressContentStreams = false }.EnableTaggedPdfCatalogMarkers());
        for (int index = 0; index < pages.Length; index++) {
            if (index > 0) builder = builder.PageBreak();
            builder = builder.Canvas(pages[index]);
        }
        return builder.ToBytes();
    }

    private static byte[] RenumberLastContentMcid(byte[] bytes) {
        string text = PdfEncoding.Latin1GetString(bytes);
        Match last = Regex.Matches(text, @"/MCID (\d)\s*>>\s*BDC").Cast<Match>().Last();
        Assert.True(last.Success);
        char[] chars = text.ToCharArray();
        chars[last.Groups[1].Index] = '9';
        return PdfEncoding.Latin1GetBytes(new string(chars));
    }
}
