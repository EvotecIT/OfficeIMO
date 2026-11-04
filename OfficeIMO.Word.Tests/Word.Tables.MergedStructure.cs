using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordTableMergedStructureTests {
    private static WordTable CreateTable(WordDocument document, int rows, int columns) {
        WordTable table = document.AddTable(rows, columns);
        for (int r = 0; r < rows; r++) {
            for (int c = 0; c < columns; c++) table.Rows[r].Cells[c].AddParagraph("R" + r + "C" + c, true);
        }
        return table;
    }

    private static string Text(WordTableCell cell) => string.Join("|", cell.Paragraphs.Select(paragraph => paragraph.Text).Where(text => text.Length > 0));

    private static byte[] Save(WordDocument document) {
        using var output = new MemoryStream();
        document.Save(output);
        return output.ToArray();
    }

    /// <summary>Replaces the first three cells of a row with one that spans them, as other producers write it.</summary>
    private static WordDocument LoadWithSpanningRow(int rowIndex) {
        byte[] bytes;
        using (WordDocument document = WordDocument.Create()) {
            CreateTable(document, 3, 3);
            bytes = Save(document);
        }
        var stream = new MemoryStream();
        stream.Write(bytes, 0, bytes.Length);
        using (WordprocessingDocument package = WordprocessingDocument.Open(stream, true)) {
            TableRow row = package.MainDocumentPart!.Document!.Body!.Elements<Table>().Single().Elements<TableRow>().ElementAt(rowIndex);
            TableCell[] cells = row.Elements<TableCell>().ToArray();
            cells[1].Remove();
            cells[2].Remove();
            cells[0].TableCellProperties!.GridSpan = new GridSpan { Val = 3 };
        }
        stream.Position = 0;
        return WordDocument.Load(stream);
    }

    [Fact]
    public void RowsAddedAfterASpanningRowCoverEveryGridColumn() {
        using WordDocument document = LoadWithSpanningRow(2);
        WordTable table = document.Tables.Single();

        WordTableRow added = table.AddRow();

        Assert.Equal(3, added.CellsCount);
    }

    [Fact]
    public void RowsInsertedBelowASpanningHeaderTakeAValueForEveryGridColumn() {
        using WordDocument document = LoadWithSpanningRow(0);
        WordTable table = document.Tables.Single();

        WordTableRow row = table.InsertRow(1, "A", "B", "C");

        Assert.Equal(new[] { "A", "B", "C" }, row.Cells.Select(Text));
    }

    [Fact]
    public void AColumnInsertedThroughASpanningCellWidensIt() {
        using WordDocument document = LoadWithSpanningRow(0);
        WordTable table = document.Tables.Single();

        IReadOnlyList<WordTableCell> added = table.InsertColumn(1, "wide", "B", "C");

        TableRow header = table.Rows[0]._tableRow;
        Assert.Equal(4, header.Elements<TableCell>().Single().TableCellProperties!.GridSpan!.Val!.Value);
        Assert.Equal("R0C0|wide", Text(added[0]));
        Assert.Equal(new[] { "R1C0", "B", "R1C1", "R1C2" }, table.Rows[1].Cells.Select(Text));
    }

    [Fact]
    public void ARowInsertedInsideAVerticalMergeJoinsIt() {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateTable(document, 3, 2);
        table.MergeCells(0, 0, 3, 1);

        table.InsertRow(1, "kept", "plain");

        Assert.Equal(WordCellMerge.Continue, table.Rows[1].Cells[0].VerticalMerge);
        Assert.Equal("R0C0|kept", Text(table.Rows[0].Cells[0]));
        Assert.Equal("plain", Text(table.Rows[1].Cells[1]));
    }

    [Fact]
    public void AColumnInsertedInsideAHorizontalMergeJoinsIt() {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateTable(document, 2, 3);
        table.MergeCells(0, 0, 1, 3);

        table.InsertColumn(1, "kept", "plain");

        Assert.Equal(WordCellMerge.Continue, table.Rows[0].Cells[1].HorizontalMerge);
        Assert.Equal("R0C0|kept", Text(table.Rows[0].Cells[0]));
        Assert.Equal("plain", Text(table.Rows[1].Cells[1]));
    }

    [Fact]
    public void RemovingTheFirstRowOfAMergeHandsTheCellToTheNextRow() {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateTable(document, 3, 2);
        table.MergeCells(0, 0, 3, 1);

        table.Rows[0].Remove();

        Assert.Equal(WordCellMerge.Restart, table.Rows[0].Cells[0].VerticalMerge);
        Assert.Equal("R0C0", Text(table.Rows[0].Cells[0]));
        Assert.Equal(WordCellMerge.Continue, table.Rows[1].Cells[0].VerticalMerge);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void CopiedTablesAndRowsLeaveTheirBookmarksOnTheOriginal(int copies) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateTable(document, 1, 1);
        table.Rows[0].Cells[0].Paragraphs[0].AddBookmark("Target");

        table.CopyRow(table.Rows[0]);
        for (int copy = 1; copy < copies; copy++) table.Clone();

        using var package = WordprocessingDocument.Open(new MemoryStream(Save(document)), false);
        Assert.Single(package.MainDocumentPart!.Document!.Descendants<BookmarkStart>());
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(package));
    }

    [Fact]
    public void EmptiedTablesRowsAndCellsAreNotSaved() {
        using WordDocument document = WordDocument.Create();
        CreateTable(document, 1, 1).Rows[0].Remove();
        WordTable cellless = CreateTable(document, 2, 1);
        cellless.Rows[1].Cells[0].Remove();
        WordTable paragraphless = CreateTable(document, 1, 1);
        foreach (WordParagraph paragraph in paragraphless.Rows[0].Cells[0].Paragraphs.ToList()) paragraph.Remove();

        using var package = WordprocessingDocument.Open(new MemoryStream(Save(document)), false);
        Table[] tables = package.MainDocumentPart!.Document!.Body!.Elements<Table>().ToArray();
        Assert.Equal(2, tables.Length);
        Assert.Single(tables[0].Elements<TableRow>());
        Assert.All(tables.SelectMany(table => table.Descendants<TableCell>()),
            cell => Assert.IsType<Paragraph>(cell.ChildElements.Last()));
    }

    [Fact]
    public void RemovedContentTakesItsImagesAndChartsWithIt() {
        string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "Kulek.jpg");
        using WordDocument document = WordDocument.Create();
        WordTable removedTable = CreateTable(document, 1, 1);
        removedTable.Rows[0].Cells[0].Paragraphs[0].AddImage(imagePath);
        document.AddChart(OfficeChartType.Column, new OfficeChartData(new[] { "a" }, new[] { new OfficeSeries("S", new double[] { 1 }) }));
        WordParagraph chartParagraph = document.Paragraphs.Last(paragraph => paragraph.IsChart);

        removedTable.Remove();
        chartParagraph.Remove();

        using var package = WordprocessingDocument.Open(new MemoryStream(Save(document)), false);
        Assert.Empty(package.MainDocumentPart!.ImageParts);
        Assert.Empty(package.MainDocumentPart.ChartParts);
    }

    [Fact]
    public void ContentSharingAnImageKeepsItWhenOneCopyIsRemoved() {
        string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "Kulek.jpg");
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateTable(document, 1, 1);
        table.Rows[0].Cells[0].Paragraphs[0].AddImage(imagePath);
        WordTable copy = table.Clone();

        copy.Remove();

        using var package = WordprocessingDocument.Open(new MemoryStream(Save(document)), false);
        Assert.Single(package.MainDocumentPart!.ImageParts);
    }
}
