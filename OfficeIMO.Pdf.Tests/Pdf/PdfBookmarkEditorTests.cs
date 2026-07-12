using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfBookmarkEditorTests {
    [Fact]
    public void Edit_AddsRenamesMovesNestsRetargetsRemovesAndRebuilds() {
        byte[] source = PdfDocument.Create(new PdfOptions { CreateOutlineFromHeadings = true })
            .H1("First")
            .Paragraph(p => p.Text("Page one"))
            .PageBreak()
            .H1("Second")
            .H2("Detail")
            .Paragraph(p => p.Text("Page two"))
            .ToBytes();

        PdfBookmarkEditResult edited = PdfDocument.Open(source).Bookmarks.Edit(session => {
            PdfBookmarkNode first = session.Roots[0];
            PdfBookmarkNode second = session.Roots[1];
            session.Rename(first.Id, "Renamed first");
            PdfBookmarkNode added = session.Add("Added", 1, destinationTop: 500);
            session.Move(added.Id, second.Id, 0);
            session.Retarget(second.Id, 1, 400);
            session.Remove(second.Children.Last().Id);
        });

        Assert.Equal(PdfMutationExecutionMode.FullRewrite, edited.MutationPlan.ExecutionMode);
        Assert.Equal("Renamed first", edited.Outlines[0].Title);
        Assert.Equal(1, edited.Outlines[1].PageNumber);
        Assert.Equal("Added", Assert.Single(edited.Outlines[1].Children).Title);
        Assert.Empty(edited.OpenDocument().Bookmarks.Validate());

        PdfBookmarkEditResult rebuilt = edited.OpenDocument().Bookmarks.Edit(session => session.RebuildFromHeadings());
        Assert.Contains(rebuilt.Outlines, static outline => outline.Title == "First");
        Assert.Contains(rebuilt.Outlines, static outline => outline.Title == "Second");
        Assert.Contains(rebuilt.Outlines.SelectMany(static outline => outline.Children), static outline => outline.Title == "Detail");
    }
}
