using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherTableMergingTests {
    [Fact]
    public void Authored_merges_retain_styled_paragraphs_tracks_and_complete_anchor_ownership() {
        PublisherDocument document = Create();
        uint id = document.Pages[0].Tables[0].Id;
        for (int row = 0; row < 3; row++)
            for (int column = 0; column < 4; column++) {
                int index = row * 4 + column;
                document.ReplaceTableCellParagraphs(id, row, column, new[] { Paragraph(((char)('A' + index)).ToString(), 10 + index, index % 2 == 0) });
            }
        PublisherTable snapshot = document.Pages[0].Tables[0];
        document.MergeTableCells(id, 0, 0, 1, 2);
        document.MergeTableCells(id, 1, 0, 2, 1);
        document.MergeTableCells(id, 1, 2, 2, 2);
        PublisherTable table = document.Pages[0].Tables[0];
        Assert.Equal(12, snapshot.Cells.Count);
        Assert.Equal(7, table.Cells.Count);
        Assert.Equal(new[] { 90D, 110D, 130D, 150D }, table.ColumnWidths);
        Assert.Equal(new[] { 42D, 54D, 66D }, table.RowHeights);
        Cell(0, 0, 1, 2, 200, 42, "A\nB");
        Cell(1, 0, 2, 1, 90, 120, "E\nI");
        Cell(1, 2, 2, 2, 280, 120, "G\nH\nK\nL");
        Cell(0, 2, 1, 1, 130, 42, "C");
        Cell(0, 3, 1, 1, 150, 42, "D");
        Cell(1, 1, 1, 1, 110, 54, "F");
        Cell(2, 1, 1, 1, 110, 66, "J");
        Assert.Equal("A\nB\nC\nD\nE\nI\nF\nG\nH\nK\nL\nJ\n", document.TextStories.Single().Text);
        foreach (PublisherTableCell cell in table.Cells)
            foreach (OfficeRichTextParagraph paragraph in cell.Paragraphs) {
                OfficeRichTextRun run = Assert.Single(paragraph.Runs);
                int index = run.Text[0] - 'A';
                Assert.Equal(10D + index, run.FontSize);
                Assert.Equal(index % 2 == 0, run.Bold);
            }
        string svg = document.ToSvg();
        document.MergeTableCells(id, 1, 2, 2, 2);
        Assert.Equal(svg, document.ToSvg());
        Assert.Throws<ArgumentException>(() => document.ReplaceTableCellParagraphs(id, 2, 3, new[] { Paragraph("Covered") }));
        Assert.Equal(svg, document.ToSvg());
        document.ReplaceTableCellParagraphs(id, 1, 2, new[] { Paragraph("Edited merged anchor") });
        Assert.Equal("Edited merged anchor", document.Pages[0].Tables[0].Cells.Single(cell => cell.RowIndex == 1 && cell.ColumnIndex == 2).Text);
        Assert.Equal("J", document.Pages[0].Tables[0].Cells.Single(cell => cell.RowIndex == 2 && cell.ColumnIndex == 1).Text);
        document.MergeTableCells(id, 0, 0, 3, 4);
        PublisherTableCell complete = Assert.Single(document.Pages[0].Tables[0].Cells);
        Assert.Equal((3, 4, 480D, 162D), (complete.RowSpan, complete.ColumnSpan, complete.Width, complete.Height));
        Assert.Equal("A\nB\nC\nD\nE\nI\nF\nEdited merged anchor\nJ", complete.Text);

        void Cell(int row, int column, int rows, int columns, double width, double height, string text) {
            PublisherTableCell cell = Assert.Single(table.Cells, item => item.RowIndex == row && item.ColumnIndex == column);
            Assert.Equal((rows, columns, width, height, text), (cell.RowSpan, cell.ColumnSpan, cell.Width, cell.Height, cell.Text));
        }
    }

    [Theory]
    [InlineData(-1, 0, 1, 1)]
    [InlineData(0, -1, 1, 1)]
    [InlineData(0, 0, 0, 1)]
    [InlineData(0, 0, 1, 0)]
    [InlineData(0, 0, int.MaxValue, 1)]
    [InlineData(0, 0, 1, int.MaxValue)]
    [InlineData(0, 1, 2, 1)]
    public void Invalid_or_partial_merge_ranges_leave_the_authored_publication_unchanged(int row, int column, int rows, int columns) {
        PublisherDocument document = Create();
        uint id = document.Pages[0].Tables[0].Id;
        document.MergeTableCells(id, 0, 0, 1, 2);
        string before = document.ToSvg();
        Assert.ThrowsAny<ArgumentException>(() => document.MergeTableCells(id, row, column, rows, columns));
        Assert.Equal(before, document.ToSvg());
        Assert.Equal(11, document.Pages[0].Tables[0].Cells.Count);
    }

    [Fact]
    public void Imported_topology_and_cancelled_authoring_are_refused_without_mutation() {
        PublisherDocument imported = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/TableCreation/native-tablecreation.pub"));
        uint id = imported.Pages[0].Tables[0].Id;
        string nativeBefore = imported.ToSvg();
        Assert.Throws<NotSupportedException>(() => imported.MergeTableCells(id, 0, 0, 1, 2));
        Assert.Equal(nativeBefore, imported.ToSvg());
        PublisherDocument authored = Create();
        string before = authored.ToSvg();
        Assert.Throws<OperationCanceledException>(() => authored.MergeTableCells(authored.Pages[0].Tables[0].Id, 0, 0, 1, 2, new CancellationToken(true)));
        Assert.Equal(before, authored.ToSvg());
        Assert.Equal(12, authored.Pages[0].Tables[0].Cells.Count);
    }

    private static PublisherDocument Create() {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        document.AddTable(document.Pages[0].Id, 72, 144, new[] { 90D, 110D, 130D, 150D }, new[] { 42D, 54D, 66D });
        return document;
    }
    private static OfficeRichTextParagraph Paragraph(string text, double size = 12, bool bold = false) =>
        new(new[] { new OfficeRichTextRun(text, size, OfficeColor.Black, bold, fontFamily: "Arial") }, lineHeight: 18);
}
