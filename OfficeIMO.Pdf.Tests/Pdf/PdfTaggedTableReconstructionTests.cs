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
    public void TaggedCellSourceLinesAreTheLinesHoldingItsText() {
        // Cells far enough apart that each one is its own page line: a cell's source lines are the lines its text
        // sits on, not every line on the same baseline.
        byte[] bytes = CreateTaggedDocument(canvas => canvas.Structure(PdfCanvasStructureRole.Table, table => table
            .Structure(PdfCanvasStructureRole.TableBody, body => {
                string[][] rows = {
                    new[] { "Market", "Spot", "Futures", "Options" },
                    new[] { "Power", "Yes", "No", "Yes" },
                    new[] { "Gas", "No", "Yes", "No" }
                };
                for (int rowIndex = 0; rowIndex < rows.Length; rowIndex++) {
                    string[] cells = rows[rowIndex];
                    double top = 120D + (rowIndex * 20D);
                    body.Structure(PdfCanvasStructureRole.TableRow, row => {
                        for (int column = 0; column < cells.Length; column++) {
                            string text = cells[column];
                            row.Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, column, top, text));
                        }
                    });
                }
            })));

        using PdfDocument document = PdfDocument.Load(new MemoryStream(bytes));
        PdfDocumentReadResult read = document.Read();

        PdfUnderstandingTableCandidate table = Assert.Single(read.Pages[0].Analysis.TableCandidates,
            static candidate => candidate.DetectionKind == "tagged-structure");
        Assert.Equal(12, table.SourceLines.Count);
        Assert.Equal(
            table.Rows.SelectMany(static row => row).OrderBy(static text => text, StringComparer.Ordinal),
            table.SourceLines.Select(static line => line.Text).OrderBy(static text => text, StringComparer.Ordinal));
    }

    [Fact]
    public void TaggedCellTextOnTwoPageLinesKeepsEachLinesOwnRuns() {
        // One cell paints two pieces on the same baseline, far enough apart to be two page lines.
        byte[] bytes = CreateTaggedDocument(canvas => canvas.Structure(PdfCanvasStructureRole.Table, table => table
            .Structure(PdfCanvasStructureRole.TableBody, body => {
                body.Structure(PdfCanvasStructureRole.TableRow, row => row
                    .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, 0, 120D, "Region"))
                    .Structure(PdfCanvasStructureRole.TableCell, cell => {
                        Lines(cell, 1, 120D, "North");
                        Lines(cell, 3, 120D, "Coast");
                    }));
                body.Structure(PdfCanvasStructureRole.TableRow, row => row
                    .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, 0, 140D, "Sales"))
                    .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, 1, 140D, "12")));
            })));

        using PdfDocument document = PdfDocument.Load(new MemoryStream(bytes));
        PdfDocumentReadResult read = document.Read();

        PdfUnderstandingTableCandidate table = Assert.Single(read.Pages[0].Analysis.TableCandidates,
            static candidate => candidate.DetectionKind == "tagged-structure");
        Assert.Equal("North Coast", table.Rows[0][1]);
        PdfUnderstandingLine[] split = table.SourceLines
            .Where(static line => line.Text.Contains("North", StringComparison.Ordinal) || line.Text.Contains("Coast", StringComparison.Ordinal))
            .ToArray();
        Assert.Equal(new[] { "Coast", "North" }, split.Select(static line => line.Text).OrderBy(static text => text, StringComparer.Ordinal));
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
            canvas => TablePage(canvas, new[] { ("Gamma", "3") }, header: false));
        using PdfDocument document = PdfDocument.Load(new MemoryStream(bytes));

        PdfDocumentReadResult read = document.Read(new PdfLoadOptions { PageSelection = PdfPageSelection.From(2) });

        // One body row is enough on the selected page: the rest of the table is on page 1.
        PdfLogicalTable table = Assert.Single(Assert.Single(read.Pages).Tables);
        Assert.Equal("tagged-structure", table.DetectionKind);
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);
        Assert.Equal(new[] { "Item", "Count" }, data.Columns);
        Assert.Equal(new[] { "Gamma|3" }, data.Rows.Select(static row => string.Join("|", row)));
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

    [Fact]
    public void InterleavedTaggedTablesOnTheSamePagesEachStayWhole() {
        // Two Table elements both continue from page 1 to page 2, so their segments alternate in reading order.
        byte[] bytes = CreateTaggedDocument(
            canvas => {
                TablePage(canvas, new[] { ("Alpha", "1"), ("Beta", "2") }, key: "a", top: 100D);
                TablePage(canvas, new[] { ("North", "7"), ("South", "8") }, key: "b", top: 400D);
            },
            canvas => {
                TablePage(canvas, new[] { ("Gamma", "3") }, header: false, key: "a", top: 100D);
                TablePage(canvas, new[] { ("East", "9") }, header: false, key: "b", top: 400D);
            });
        using PdfDocument document = PdfDocument.Load(new MemoryStream(bytes));

        PdfDocumentReadResult read = document.Read();

        Assert.Collection(read.GetTableContinuationGroups(),
            group => {
                Assert.Equal(new[] { 1, 2 }, group.Segments.Select(static segment => segment.PageNumber));
                Assert.Equal(new[] { "Alpha|1", "Beta|2", "Gamma|3" }, group.Rows.Select(static row => string.Join("|", row)));
            },
            group => {
                Assert.Equal(new[] { 1, 2 }, group.Segments.Select(static segment => segment.PageNumber));
                Assert.Equal(new[] { "North|7", "South|8", "East|9" }, group.Rows.Select(static row => string.Join("|", row)));
            });
    }

    [Fact]
    public void RowSpanStartingInAnEmptyRowStaysMergedOverTheRowsItCovers() {
        byte[] bytes = CreateTaggedDocument(canvas => canvas.Structure(PdfCanvasStructureRole.Table, table => table
            .Structure(PdfCanvasStructureRole.TableBody, head => head
                .Structure(PdfCanvasStructureRole.TableRow, row => row
                    .Structure(PdfCanvasStructureRole.TableHeaderCell, cell => Lines(cell, 0, 100D, "Group"))
                    .Structure(PdfCanvasStructureRole.TableHeaderCell, cell => Lines(cell, 1, 100D, "Item"))
                    .Structure(PdfCanvasStructureRole.TableHeaderCell, cell => Lines(cell, 2, 100D, "Count"))),
                new PdfCanvasStructureOptions { TableRowGroupType = "THead" })
            .Structure(PdfCanvasStructureRole.TableBody, body => body
                .Structure(PdfCanvasStructureRole.TableRow, row => row
                    .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, 0, 120D, string.Empty),
                        new PdfCanvasStructureOptions { RowSpan = 3 })
                    .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, 1, 120D, string.Empty))
                    .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, 2, 120D, string.Empty)))
                .Structure(PdfCanvasStructureRole.TableRow, row => row
                    .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, 1, 140D, "Alpha"))
                    .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, 2, 140D, "1")))
                .Structure(PdfCanvasStructureRole.TableRow, row => row
                    .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, 1, 160D, "Beta"))
                    .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, 2, 160D, "2")))
                .Structure(PdfCanvasStructureRole.TableRow, row => row
                    .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, 0, 180D, "Other"))
                    .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, 1, 180D, "Gamma"))
                    .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, 2, 180D, "3"))))));
        using PdfDocument document = PdfDocument.Load(new MemoryStream(bytes));

        PdfLogicalTable table = Assert.Single(document.Read().Pages[0].Tables);

        // The empty first body row is dropped; its span now starts on the next row and covers the one after.
        Assert.Equal(new[] { "Group|Item|Count", "|Alpha|1", "|Beta|2", "Other|Gamma|3" },
            table.Rows.Select(static row => string.Join("|", row)));
        Assert.Equal(2, table.Cells.Single(cell => cell.RowIndex == 1 && cell.ColumnIndex == 0).RowSpan);
        Assert.True(table.Cells.Single(cell => cell.RowIndex == 2 && cell.ColumnIndex == 0).IsCoveredBySpan);
        Assert.False(table.Cells.Single(cell => cell.RowIndex == 3 && cell.ColumnIndex == 0).IsCoveredBySpan);
    }

    [Fact]
    public void SelectedContinuationPageKeepsColumnsFilledOnOtherPages() {
        byte[] bytes = CreateTaggedDocument(
            canvas => GridPage(canvas, new[] { new[] { "Item", "Note", "Count" } },
                new[] { new[] { (0, "Alpha", 1), (1, "First", 1), (2, "1", 1) } }),
            canvas => GridPage(canvas, Array.Empty<string[]>(),
                new[] { new[] { (0, "Gamma", 1), (1, string.Empty, 1), (2, "3", 1) } }));
        using PdfDocument document = PdfDocument.Load(new MemoryStream(bytes));

        PdfDocumentReadResult read = document.Read(new PdfLoadOptions { PageSelection = PdfPageSelection.From(2) });

        // The Note column is empty on page 2 but filled on page 1, so it stays.
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(Assert.Single(Assert.Single(read.Pages).Tables));
        Assert.Equal(new[] { "Item", "Note", "Count" }, data.Columns);
        Assert.Equal(new[] { "Gamma||3" }, data.Rows.Select(static row => string.Join("|", row)));
    }

    [Fact]
    public void RepeatedMultiRowHeaderStaysOutOfTheContinuedRows() {
        string[][] header = { new[] { "Product", "Price", string.Empty }, new[] { string.Empty, "Bid", "Ask" } };
        byte[] bytes = CreateTaggedDocument(
            canvas => GridPage(canvas, header, new[] { new[] { (0, "Alpha", 1), (1, "10", 1), (2, "11", 1) } }),
            canvas => GridPage(canvas, header, new[] { new[] { (0, "Beta", 1), (1, "20", 1), (2, "21", 1) } }));
        using PdfDocument document = PdfDocument.Load(new MemoryStream(bytes));

        PdfLogicalTableContinuationGroup group = Assert.Single(document.Read().GetTableContinuationGroups());

        Assert.Equal(new[] { "Product", "Price Bid", "Ask" }, group.Columns);
        Assert.Equal(new[] { "Alpha|10|11", "Beta|20|21" }, group.Rows.Select(static row => string.Join("|", row)));
    }

    [Fact]
    public void RowSpanContinuingOnTheNextPageStaysMergedThere() {
        byte[] bytes = CreateTaggedDocument(
            canvas => GridPage(canvas, new[] { new[] { "Group", "Item", "Count" } },
                new[] { new[] { (0, "North", 3), (1, "Alpha", 1), (2, "1", 1) } }),
            canvas => GridPage(canvas, Array.Empty<string[]>(), new[] {
                new[] { (1, "Beta", 1), (2, "2", 1) },
                new[] { (1, "Gamma", 1), (2, "3", 1) },
                new[] { (0, "South", 1), (1, "Delta", 1), (2, "4", 1) } }));
        using PdfDocument document = PdfDocument.Load(new MemoryStream(bytes));

        PdfLogicalTable continued = Assert.Single(document.Read().Pages[1].Tables);

        Assert.Equal(new[] { "|Beta|2", "|Gamma|3", "South|Delta|4" }, continued.Rows.Select(static row => string.Join("|", row)));
        Assert.Equal(2, continued.Cells.Single(cell => cell.RowIndex == 0 && cell.ColumnIndex == 0).RowSpan);
        Assert.True(continued.Cells.Single(cell => cell.RowIndex == 1 && cell.ColumnIndex == 0).IsCoveredBySpan);
    }

    [Fact]
    public void WordExportPlacesBodySpansBelowAStackedHeader() {
        byte[] bytes = CreateTaggedDocument(canvas => GridPage(canvas,
            new[] { new[] { "Product", "Price", string.Empty }, new[] { string.Empty, "Bid", "Ask" } },
            new[] {
                new[] { (0, "North", 2), (1, "10", 1), (2, "11", 1) },
                new[] { (1, "20", 1), (2, "21", 1) } }));
        using PdfDocument document = PdfDocument.Load(new MemoryStream(bytes));

        using OfficeIMO.Word.WordDocument word = document.Read().ConvertToWordWithEvidence().Value;
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(word.ToBytes()), false);

        // The two header rows become one Word header row; the span starts on the first body row.
        TableRow[] rows = Assert.Single(package.MainDocumentPart!.Document!.Body!.Elements<Table>()).Elements<TableRow>().ToArray();
        Assert.Equal(3, rows.Length);
        Assert.Equal("Price Bid", rows[0].Elements<TableCell>().ElementAt(1).InnerText);
        Assert.Equal(MergedCellValues.Restart, rows[1].Elements<TableCell>().First().TableCellProperties!.VerticalMerge!.Val!.Value);
        Assert.Equal("North", rows[1].Elements<TableCell>().First().InnerText);
        Assert.NotNull(rows[2].Elements<TableCell>().First().TableCellProperties!.VerticalMerge);
        Assert.Equal(new[] { "20", "21" }, rows[2].Elements<TableCell>().Skip(1).Select(static cell => cell.InnerText));
    }

    [Fact]
    public void StyledFirstRowOfATaggedContinuationStaysData() {
        byte[] bytes = CreateTaggedDocument(
            canvas => TablePage(canvas, ("Alpha", "1"), ("Beta", "2")),
            canvas => TablePage(canvas, new[] { ("Gamma", "3"), ("Delta", "4") }, header: false));
        using PdfDocument document = PdfDocument.Load(new MemoryStream(bytes));
        PdfLogicalTable continued = Assert.Single(document.Read().Pages[1].Tables);
        // A bold first row over plain rows is what styled-header detection looks for in untagged tables.
        var bold = new PdfLogicalTableCellStyle(null, null, isBold: true, isItalic: false, fontSize: null, fontName: null);
        var plain = new PdfLogicalTableCellStyle(null, null, isBold: false, isItalic: false, fontSize: null, fontName: null);
        continued.CellStyles = new[] { new[] { bold, bold }, new[] { plain, plain } };

        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(continued);

        Assert.Equal(new[] { "Item", "Count" }, data.Columns);
        Assert.Equal(new[] { "Gamma|3", "Delta|4" }, data.Rows.Select(static row => string.Join("|", row)));
    }

    [Fact]
    public void SelectedContinuationPageDropsAColumnWithoutTextOnOtherPages() {
        byte[] bytes = RemoveTextShowing(CreateTaggedDocument(
            canvas => GridPage(canvas, new[] { new[] { "Item", "Nil", "Count" } },
                new[] { new[] { (0, "Alpha", 1), (1, "Nil", 1), (2, "1", 1) } }),
            canvas => GridPage(canvas, Array.Empty<string[]>(),
                new[] { new[] { (0, "Gamma", 1), (1, "Nil", 1), (2, "3", 1) } })), "Nil");
        using PdfDocument document = PdfDocument.Load(new MemoryStream(bytes));

        PdfLogicalTable complete = document.Read().Pages[1].Tables.Single();
        PdfLogicalTable selected = Assert.Single(Assert.Single(
            document.Read(new PdfLoadOptions { PageSelection = PdfPageSelection.From(2) }).Pages).Tables);

        // The middle column's marked content shows no text on any page, so both reads drop it.
        Assert.Equal(new[] { "Gamma|3" }, complete.Rows.Select(static row => string.Join("|", row)));
        Assert.Equal(new[] { "Gamma|3" }, selected.Rows.Select(static row => string.Join("|", row)));
    }

    [Fact]
    public void WordExportStylesBodyRowsBelowAStackedHeader() {
        byte[] bytes = CreateTaggedDocument(canvas => GridPage(canvas,
            new[] { new[] { "Product", "Price", string.Empty }, new[] { string.Empty, "Bid", "Ask" } },
            new[] {
                new[] { (0, "North", 1), (1, "10", 1), (2, "11", 1) },
                new[] { (0, "South", 1), (1, "20", 1), (2, "21", 1) } }));
        using PdfDocument document = PdfDocument.Load(new MemoryStream(bytes));
        PdfDocumentReadResult read = document.Read();
        PdfLogicalTable table = Assert.Single(read.Pages[0].Tables);
        var bold = new PdfLogicalTableCellStyle(null, null, isBold: true, isItalic: false, fontSize: null, fontName: null);
        var plain = new PdfLogicalTableCellStyle(null, null, isBold: false, isItalic: false, fontSize: null, fontName: null);
        var italic = new PdfLogicalTableCellStyle(null, null, isBold: false, isItalic: true, fontSize: null, fontName: null);
        table.CellStyles = new[] {
            new[] { bold, bold, bold }, new[] { bold, bold, bold },
            new[] { plain, plain, plain }, new[] { italic, italic, italic } };

        using OfficeIMO.Word.WordDocument word = read.ConvertToWordWithEvidence().Value;
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(word.ToBytes()), false);

        // Each body row keeps its own source row's style, not the style of the row above it.
        TableRow[] rows = Assert.Single(package.MainDocumentPart!.Document!.Body!.Elements<Table>()).Elements<TableRow>().ToArray();
        Assert.Equal(3, rows.Length);
        Run north = rows[1].Descendants<Run>().First();
        Run south = rows[2].Descendants<Run>().First();
        Assert.Equal("North", north.InnerText);
        Assert.Null(north.RunProperties?.Bold);
        Assert.Null(north.RunProperties?.Italic);
        Assert.NotNull(south.RunProperties?.Italic);
    }

    /// <summary>One page of a three-column tagged table whose Table, THead and TBody elements continue across pages.</summary>
    private static void GridPage(PdfPageCanvas canvas, string[][] headerRows, (int Column, string Text, int RowSpan)[][] bodyRows) {
        canvas.Structure(PdfCanvasStructureRole.Table, table => {
            if (headerRows.Length > 0) table.Structure(PdfCanvasStructureRole.TableBody, head => {
                for (int rowIndex = 0; rowIndex < headerRows.Length; rowIndex++) {
                    string[] texts = headerRows[rowIndex];
                    double y = 100D + (rowIndex * 20D);
                    int headerRow = rowIndex;
                    head.Structure(PdfCanvasStructureRole.TableRow, row => {
                        for (int column = 0; column < texts.Length; column++) {
                            int at = column;
                            row.Structure(PdfCanvasStructureRole.TableHeaderCell, cell => Lines(cell, at, y, texts[at]),
                                new PdfCanvasStructureOptions { StructureElementKey = "grid-h" + headerRow + "-" + at });
                        }
                    }, new PdfCanvasStructureOptions { StructureElementKey = "grid-head-" + headerRow });
                }
            }, new PdfCanvasStructureOptions { StructureElementKey = "grid-head", TableRowGroupType = "THead" });
            table.Structure(PdfCanvasStructureRole.TableBody, body => {
                for (int rowIndex = 0; rowIndex < bodyRows.Length; rowIndex++) {
                    (int Column, string Text, int RowSpan)[] cells = bodyRows[rowIndex];
                    double y = 160D + (rowIndex * 20D);
                    body.Structure(PdfCanvasStructureRole.TableRow, row => {
                        foreach ((int column, string text, int rowSpan) in cells) {
                            row.Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, column, y, text),
                                rowSpan > 1 ? new PdfCanvasStructureOptions { RowSpan = rowSpan } : null);
                        }
                    });
                }
            }, new PdfCanvasStructureOptions { StructureElementKey = "grid-body" });
        }, new PdfCanvasStructureOptions { StructureElementKey = "grid" });
    }

    private static void TablePage(PdfPageCanvas canvas, params (string Item, string Count)[] rows) =>
        TablePage(canvas, rows, swappedRow: -1);

    private static void TablePage(PdfPageCanvas canvas, (string Item, string Count)[] rows, int swappedRow = -1,
        string? strayParagraph = null, bool header = true, string key = "table", double top = 300D) {
        var tableKey = new PdfCanvasStructureOptions { StructureElementKey = key };
        var headKey = new PdfCanvasStructureOptions { StructureElementKey = key + "-head", TableRowGroupType = "THead" };
        var headRowKey = new PdfCanvasStructureOptions { StructureElementKey = key + "-head-row" };
        var bodyKey = new PdfCanvasStructureOptions { StructureElementKey = key + "-body" };
        canvas.Structure(PdfCanvasStructureRole.Table, table => {
            if (header) table.Structure(PdfCanvasStructureRole.TableBody, head => head
                .Structure(PdfCanvasStructureRole.TableRow, row => row
                    .Structure(PdfCanvasStructureRole.TableHeaderCell, cell => Lines(cell, 0, top, "Item"),
                        new PdfCanvasStructureOptions { StructureElementKey = key + "-item" })
                    .Structure(PdfCanvasStructureRole.TableHeaderCell, cell => Lines(cell, 1, top, "Count"),
                        new PdfCanvasStructureOptions { StructureElementKey = key + "-count" }), headRowKey), headKey);
            table.Structure(PdfCanvasStructureRole.TableBody, body => {
                for (int index = 0; index < rows.Length; index++) {
                    (string item, string count) = rows[index];
                    double y = top + 20D + (index * 20D);
                    bool swapped = index == swappedRow;
                    body.Structure(PdfCanvasStructureRole.TableRow, row => row
                        .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, swapped ? 1 : 0, y, swapped ? count : item))
                        .Structure(PdfCanvasStructureRole.TableCell, cell => Lines(cell, swapped ? 0 : 1, y, swapped ? item : count)));
                }
            }, bodyKey);
            if (strayParagraph is not null) {
                table.Structure(PdfCanvasStructureRole.Paragraph, paragraph => Lines(paragraph, 0, top + 200D, strayParagraph));
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

    /// <summary>Comments out the operators that show <paramref name="text"/>, leaving its marked-content sequences textless.</summary>
    private static byte[] RemoveTextShowing(byte[] bytes, string text) {
        string hex = "<" + string.Concat(text.Select(static ch => ((int)ch).ToString("X2", System.Globalization.CultureInfo.InvariantCulture))) + "> Tj";
        string content = PdfEncoding.Latin1GetString(bytes);
        Assert.Contains(hex, content);
        // Same length, so the cross-reference offsets stay valid.
        return PdfEncoding.Latin1GetBytes(content.Replace(hex, "%" + new string(' ', hex.Length - 1)));
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
