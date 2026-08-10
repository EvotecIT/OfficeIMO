using OfficeIMO.Drawing;

namespace OfficeIMO.OneNote.Tests;

public sealed partial class OneNoteRenderingTests {
    [Fact]
    public void TableMeasurementSimulatesSequentialListOrdinalsWithoutChangingRenderState() {
        var page = new OneNotePage { PageSize = OneNotePageSize.Letter };
        var table = new OneNoteTable {
            BordersVisible = true,
            Layout = new OneNoteLayout { X = 0.5D, Y = 0.5D, Width = 0.9D }
        };
        table.ColumnWidths.Add(0.9D);
        var row = new OneNoteTableRow();
        var cell = new OneNoteTableCell();
        var first = new OneNoteParagraph {
            List = new OneNoteListInfo { Ordered = true, Restart = true, DisplayIndex = 9 }
        };
        first.Runs.Add(new OneNoteTextRun { Text = "X" });
        var second = new OneNoteParagraph {
            List = new OneNoteListInfo { Ordered = true }
        };
        second.Runs.Add(new OneNoteTextRun { Text = "X" });
        cell.Content.Add(first);
        cell.Content.Add(second);
        row.Cells.Add(cell);
        table.Rows.Add(row);
        page.DirectContent.Add(table);

        OfficeDrawing drawing = page.ToDrawing(new OneNotePageRenderingOptions { IncludeTitle = false });
        OfficeDrawingShape frame = Assert.Single(
            drawing.Elements.OfType<OfficeDrawingShape>(),
            shape => shape.Shape.StrokeWidth == 0.75D);
        OfficeDrawingRichText[] paragraphs = drawing.Elements.OfType<OfficeDrawingRichText>().ToArray();

        Assert.Equal(2, paragraphs.Length);
        Assert.StartsWith("9. ", paragraphs[0].Runs[0].Text, StringComparison.Ordinal);
        Assert.StartsWith("10. ", paragraphs[1].Runs[0].Text, StringComparison.Ordinal);
        Assert.True(paragraphs[1].Height > paragraphs[0].Height);
        Assert.True(paragraphs[1].Y + paragraphs[1].Height <= frame.Y + frame.Shape.Height - 4D + 0.001D);
    }

    [Theory]
    [InlineData(-20D, 0D)]
    [InlineData(0D, -20D)]
    [InlineData(20D, 0D)]
    [InlineData(0D, 20D)]
    public void FullyCulledOrderedParagraphAdvancesFollowingListNumbering(double x, double y) {
        var page = new OneNotePage { PageSize = OneNotePageSize.IndexCard };
        var culled = new OneNoteParagraph {
            Layout = new OneNoteLayout { X = x, Y = y },
            List = new OneNoteListInfo { Ordered = true, Level = 1, Restart = true, DisplayIndex = 1 }
        };
        culled.Runs.Add(new OneNoteTextRun { Text = "Culled" });
        var following = new OneNoteParagraph {
            List = new OneNoteListInfo { Ordered = true, Level = 1 }
        };
        following.Runs.Add(new OneNoteTextRun { Text = "Following" });
        page.DirectContent.Add(culled);
        page.DirectContent.Add(following);

        OfficeDrawing drawing = page.ToDrawing(new OneNotePageRenderingOptions { IncludeTitle = false });
        OfficeDrawingRichText followingText = Assert.Single(
            drawing.Elements.OfType<OfficeDrawingRichText>(),
            item => item.Runs.Any(run => run.Text.Contains("Following", StringComparison.Ordinal)));

        Assert.StartsWith("2. ", followingText.Runs[0].Text, StringComparison.Ordinal);
    }
}
