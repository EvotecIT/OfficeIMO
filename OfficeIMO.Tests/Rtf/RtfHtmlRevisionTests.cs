using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlRevisionTests {
    [Fact]
    public void RtfDocument_ToHtml_Renders_Revision_Metadata_And_RoundTrips() {
        RtfDocument document = RtfDocument.Create();
        int alice = document.AddRevisionAuthor("Alice");
        int bob = document.AddRevisionAuthor("Bob");
        RtfParagraph paragraph = document.AddParagraph("Base ");
        paragraph.AddText("Inserted")
            .SetInsertedRevision(alice, 123)
            .SetRevisionSaveIds(character: 30, insertion: 40);
        paragraph.AddText(" ");
        paragraph.AddText("Removed")
            .SetDeletedRevision(bob)
            .SetRevisionSaveIds(deletion: 50);

        string html = document.ToHtml();

        Assert.Contains("<ins data-officeimo-rtf-revision=\"inserted\" data-officeimo-rtf-revision-author-index=\"0\" data-officeimo-rtf-revision-author=\"Alice\" data-officeimo-rtf-revision-timestamp=\"123\" data-officeimo-rtf-charrsid=\"30\" data-officeimo-rtf-insrsid=\"40\">Inserted</ins>", html, StringComparison.Ordinal);
        Assert.Contains("<del data-officeimo-rtf-revision=\"deleted\" data-officeimo-rtf-revision-author-index=\"1\" data-officeimo-rtf-revision-author=\"Bob\" data-officeimo-rtf-delrsid=\"50\">Removed</del>", html, StringComparison.Ordinal);

        RtfDocument roundTrip = html.ToRtfDocumentFromHtml();
        Assert.Collection(roundTrip.RevisionAuthors,
            author => Assert.Equal("Alice", author.Name),
            author => Assert.Equal("Bob", author.Name));
        RtfParagraph roundTripParagraph = Assert.Single(roundTrip.Paragraphs);
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
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Semantic_Revision_Tags() {
        const string html = "<p>Base <ins data-officeimo-rtf-revision-author=\"Alice\" data-officeimo-rtf-revision-timestamp=\"123\" data-officeimo-rtf-insrsid=\"40\">Inserted</ins> <del data-officeimo-rtf-revision-author=\"Bob\" data-officeimo-rtf-delrsid=\"50\">Removed</del></p>";

        RtfDocument document = html.ToRtfDocumentFromHtml();

        Assert.Collection(document.RevisionAuthors,
            author => Assert.Equal("Alice", author.Name),
            author => Assert.Equal("Bob", author.Name));
        RtfParagraph paragraph = Assert.Single(document.Paragraphs);
        RtfRun inserted = Assert.Single(paragraph.Runs, run => run.Text == "Inserted");
        Assert.Equal(RtfRevisionKind.Inserted, inserted.RevisionKind);
        Assert.Equal(0, inserted.RevisionAuthorIndex);
        Assert.Equal(123, inserted.RevisionTimestampValue);
        Assert.Equal(40, inserted.InsertionRevisionSaveId);

        RtfRun removed = Assert.Single(paragraph.Runs, run => run.Text == "Removed");
        Assert.Equal(RtfRevisionKind.Deleted, removed.RevisionKind);
        Assert.Equal(1, removed.RevisionAuthorIndex);
        Assert.Equal(50, removed.DeletionRevisionSaveId);
        Assert.False(removed.Strike);

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        Assert.Contains(@"{\*\revtbl{Alice;}{Bob;}}", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\revised \revauth0 \revdttm123 \insrsid40 Inserted", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\deleted \revauth1 \delrsid50 Removed", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Revision_Save_Id_Metadata_Without_Revision_Kind() {
        const string html = "<p>Base <span data-officeimo-rtf-revision=\"none\" data-officeimo-rtf-charrsid=\"30\" data-officeimo-rtf-insrsid=\"40\" data-officeimo-rtf-delrsid=\"50\">Tracked</span></p>";

        RtfDocument document = html.ToRtfDocumentFromHtml();

        RtfRun tracked = Assert.Single(Assert.Single(document.Paragraphs).Runs, run => run.Text == "Tracked");
        Assert.Equal(RtfRevisionKind.None, tracked.RevisionKind);
        Assert.Equal(30, tracked.CharacterRevisionSaveId);
        Assert.Equal(40, tracked.InsertionRevisionSaveId);
        Assert.Equal(50, tracked.DeletionRevisionSaveId);

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        Assert.Contains(@"\charrsid30 \insrsid40 \delrsid50 Tracked", rtf, StringComparison.Ordinal);
    }
}
