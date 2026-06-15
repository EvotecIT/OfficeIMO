using OfficeIMO.Rtf;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlRevisionTableTests {
    [Fact]
    public void RtfDocument_ToHtml_RoundTrips_Revision_Tables_And_Save_Ids() {
        RtfDocument document = RtfDocument.Create();
        int alice = document.AddRevisionAuthor("Alice");
        int bob = document.AddRevisionAuthor("Bob");
        document.SetRevisionRootSaveId(7)
            .AddRevisionSaveId(15)
            .AddRevisionSaveId(1024)
            .AddRevisionSaveId(65535);
        RtfParagraph paragraph = document.AddParagraph("Base ");
        paragraph.SetRevisionSaveId(20);
        paragraph.AddText("Inserted")
            .SetInsertedRevision(alice, 123)
            .SetRevisionSaveIds(character: 30, insertion: 40);
        paragraph.AddText(" ");
        paragraph.AddText("Removed")
            .SetDeletedRevision(bob)
            .SetRevisionSaveIds(deletion: 50);

        string html = document.ToHtml(new RtfToHtmlOptions {
            FragmentOnly = false,
            NewLine = "\n"
        });

        Assert.Contains("<meta name=\"officeimo-rtf-revision-tables\" content=\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-rtf-pararsid=\"20\"", html, StringComparison.Ordinal);

        RtfDocument roundTrip = html.LoadFromHtml();

        Assert.Collection(roundTrip.RevisionAuthors,
            author => Assert.Equal("Alice", author.Name),
            author => Assert.Equal("Bob", author.Name));
        Assert.Equal(7, roundTrip.RevisionRootSaveId);
        Assert.Equal(new[] { 15, 1024, 65535 }, roundTrip.RevisionSaveIds);

        RtfParagraph roundTripParagraph = Assert.Single(roundTrip.Paragraphs);
        Assert.Equal(20, roundTripParagraph.RevisionSaveId);
        RtfRun inserted = Assert.Single(roundTripParagraph.Runs, run => run.Text == "Inserted");
        Assert.Equal(RtfRevisionKind.Inserted, inserted.RevisionKind);
        Assert.Equal(alice, inserted.RevisionAuthorIndex);
        Assert.Equal(123, inserted.RevisionTimestampValue);
        Assert.Equal(30, inserted.CharacterRevisionSaveId);
        Assert.Equal(40, inserted.InsertionRevisionSaveId);

        RtfRun removed = Assert.Single(roundTripParagraph.Runs, run => run.Text == "Removed");
        Assert.Equal(RtfRevisionKind.Deleted, removed.RevisionKind);
        Assert.Equal(bob, removed.RevisionAuthorIndex);
        Assert.Equal(50, removed.DeletionRevisionSaveId);

        string rtf = roundTrip.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        Assert.Contains(@"{\*\revtbl{Alice;}{Bob;}}", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\*\rsidtbl\rsidroot7\rsid15\rsid1024\rsid65535}", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\pararsid20", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\revised \revauth0 \revdttm123 \charrsid30 \insrsid40 Inserted", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\deleted \revauth1 \delrsid50 Removed", rtf, StringComparison.Ordinal);
    }
}
