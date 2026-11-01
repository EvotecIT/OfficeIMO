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
}
