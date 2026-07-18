using System.Text;
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
        Assert.Empty(edited.ToDocument().Bookmarks.Validate());

        PdfBookmarkEditResult rebuilt = edited.ToDocument().Bookmarks.Edit(session => session.RebuildFromHeadings());
        Assert.Contains(rebuilt.Outlines, static outline => outline.Title == "First");
        Assert.Contains(rebuilt.Outlines, static outline => outline.Title == "Second");
        Assert.Contains(rebuilt.Outlines.SelectMany(static outline => outline.Children), static outline => outline.Title == "Detail");
    }

    [Fact]
    public void Edit_PreservesDestinationModesCoordinatesAndZoom() {
        byte[] source = BuildDestinationModesPdf();

        PdfBookmarkEditResult edited = PdfBookmarkEditor.Edit(source, session => session.Rename(session.Roots[0].Id, "Renamed FitH"));

        Assert.Equal(2, edited.Outlines.Count);
        PdfOutlineItem fitHorizontal = edited.Outlines[0];
        Assert.Equal("Renamed FitH", fitHorizontal.Title);
        Assert.Equal(PdfOpenActionDestinationMode.FitHorizontal, fitHorizontal.DestinationMode);
        Assert.Equal(144d, fitHorizontal.DestinationTop);
        PdfOutlineItem xyz = edited.Outlines[1];
        Assert.Equal(PdfOpenActionDestinationMode.Xyz, xyz.DestinationMode);
        Assert.Equal(12d, xyz.DestinationLeft);
        Assert.Equal(100d, xyz.DestinationTop);
        Assert.Equal(1.5d, xyz.DestinationZoom);
    }

    private static byte[] BuildDestinationModesPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /PageMode /UseOutlines >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", string.Empty, "endstream", "endobj",
            "5 0 obj", "<< /Type /Outlines /First 6 0 R /Last 7 0 R /Count 2 >>", "endobj",
            "6 0 obj", "<< /Title (Fit horizontal) /Parent 5 0 R /Dest [3 0 R /FitH 144] /Next 7 0 R >>", "endobj",
            "7 0 obj", "<< /Title (XYZ) /Parent 5 0 R /Dest [3 0 R /XYZ 12 100 1.5] /Prev 6 0 R >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 8 >>", "%%EOF"
        }) + "\n";
        return Encoding.ASCII.GetBytes(pdf);
    }
}
