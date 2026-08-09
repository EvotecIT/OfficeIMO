using OfficeIMO.Drawing;

namespace OfficeIMO.OneNote.Tests;

public sealed partial class OneNoteRenderingTests {
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
