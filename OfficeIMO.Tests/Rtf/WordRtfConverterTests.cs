using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using System.Linq;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class WordRtfConverterTests {
    [Fact]
    public void Word_To_Rtf_Uses_Core_Model_And_RoundTrips_Text_Formatting_And_Hyperlinks() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph();
        paragraph.AddText("Hello ");
        paragraph.AddText("RTF").SetBold();
        paragraph.AddText(" at ");
        paragraph.AddHyperLink("OfficeIMO", new Uri("https://github.com/EvotecIT/OfficeIMO"), addStyle: true);

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        using WordDocument roundTrip = rtf.LoadFromRtf();

        RtfParagraph rtfParagraph = Assert.Single(rtfDocument.Paragraphs);
        Assert.Equal("Hello RTF at OfficeIMO", rtfParagraph.ToPlainText());
        Assert.Contains(rtfParagraph.Runs, run => run.Text == "RTF" && run.Bold);
        Assert.Contains(rtfParagraph.Runs, run => run.Text == "OfficeIMO" && run.Hyperlink != null);
        Assert.Contains(@"HYPERLINK ""https://github.com/EvotecIT/OfficeIMO""", rtf, StringComparison.Ordinal);
        Assert.Equal("Hello RTF at ", string.Concat(roundTrip.Paragraphs.Select(paragraphFromRtf => paragraphFromRtf.Text)));
        WordHyperLink roundTripLink = Assert.Single(roundTrip.HyperLinks);
        Assert.Equal("OfficeIMO", roundTripLink.Text);
        Assert.Equal("https://github.com/EvotecIT/OfficeIMO", roundTripLink.Uri?.ToString());
    }

    [Fact]
    public void Word_Rtf_Bridge_Carries_Insert_And_Delete_Revisions() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph("Before ");
        paragraph.AddInsertedText("Added", "Alice", new DateTime(2026, 1, 1));
        paragraph.AddDeletedText("Removed", "Bob", new DateTime(2026, 1, 2));

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        using WordDocument roundTrip = rtf.LoadFromRtf();

        Assert.Collection(rtfDocument.RevisionAuthors,
            author => Assert.Equal("Alice", author.Name),
            author => Assert.Equal("Bob", author.Name));
        RtfParagraph rtfParagraph = Assert.Single(rtfDocument.Paragraphs);
        Assert.Contains(rtfParagraph.Runs, run => run.Text == "Added" && run.RevisionKind == RtfRevisionKind.Inserted && run.RevisionAuthorIndex == 0);
        Assert.Contains(rtfParagraph.Runs, run => run.Text == "Removed" && run.RevisionKind == RtfRevisionKind.Deleted && run.RevisionAuthorIndex == 1);
        Assert.Contains(@"{\*\revtbl{Alice;}{Bob;}}", rtf, StringComparison.Ordinal);
        Assert.Contains(roundTrip._document.Body!.Descendants<InsertedRun>(), run => run.Author?.Value == "Alice" && run.InnerText == "Added");
        Assert.Contains(roundTrip._document.Body!.Descendants<DeletedRun>(), run => run.Author?.Value == "Bob" && run.InnerText == "Removed");
    }

    [Fact]
    public void Word_Rtf_Bridge_Carries_Comments_As_Annotations() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph("Comment target");
        paragraph.AddComment("Alice Reviewer", "AR", "Review note");

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        using WordDocument roundTrip = rtf.LoadFromRtf();

        RtfRun run = Assert.Single(Assert.Single(rtfDocument.Paragraphs).Runs);
        Assert.NotNull(run.Note);
        RtfNote note = run.Note!;
        Assert.Equal(RtfNoteKind.Annotation, note.Kind);
        Assert.Equal("Alice Reviewer", note.Author);
        Assert.Equal("Review note", note.ToPlainText());
        Assert.Contains(@"{\*\atnauthor Alice Reviewer}", rtf, StringComparison.Ordinal);
        WordComment comment = Assert.Single(roundTrip.Comments);
        Assert.Equal("Alice Reviewer", comment.Author);
        Assert.Equal("Review note", comment.Text);
    }

    [Fact]
    public void Word_To_Rtf_Bridge_Carries_Rich_Comment_Paragraphs_As_Annotations() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph("Comment target");
        paragraph.AddComment("Alice Reviewer", "AR", "placeholder");

        Comment comment = Assert.Single(word._wordprocessingDocument.MainDocumentPart!.WordprocessingCommentsPart!.Comments!.Elements<Comment>());
        comment.RemoveAllChildren();
        comment.Append(
            new Paragraph(
                new Run(new Text("First ")),
                new Run(new RunProperties(new Bold()), new Text("bold"))),
            new Paragraph(
                new Run(new RunProperties(new Italic()), new Text("Second"))));

        RtfDocument rtfDocument = word.ToRtfDocument();

        RtfNote note = Assert.Single(Assert.Single(rtfDocument.Paragraphs).Runs).Note!;
        Assert.Equal(RtfNoteKind.Annotation, note.Kind);
        Assert.Equal("Alice Reviewer", note.Author);
        Assert.Collection(note.Paragraphs,
            first => {
                Assert.Equal("First bold", first.ToPlainText());
                Assert.Contains(first.Runs, run => run.Text == "bold" && run.Bold);
            },
            second => {
                Assert.Equal("Second", second.ToPlainText());
                Assert.Contains(second.Runs, run => run.Text == "Second" && run.Italic);
            });
    }

    [Fact]
    public void Rtf_To_Word_Bridge_Carries_Rich_Annotations_As_Comments() {
        RtfDocument rtfDocument = RtfDocument.Create();
        RtfParagraph paragraph = rtfDocument.AddParagraph();
        RtfRun run = paragraph.AddText("Comment target");
        var note = new RtfNote(RtfNoteKind.Annotation) {
            Author = "Alice Reviewer",
            Created = new DateTime(2026, 1, 2, 3, 4, 5, DateTimeKind.Utc)
        };
        RtfParagraph first = note.AddParagraph("First ");
        first.AddText("bold").SetBold();
        RtfParagraph second = note.AddParagraph();
        second.AddText("Second").SetItalic();
        run.SetNote(note);

        using WordDocument word = rtfDocument.ToWordDocument();

        WordComment wordComment = Assert.Single(word.Comments);
        Assert.Equal("Alice Reviewer", wordComment.Author);
        Comment comment = Assert.Single(word._wordprocessingDocument.MainDocumentPart!.WordprocessingCommentsPart!.Comments!.Elements<Comment>());
        List<Paragraph> commentParagraphs = comment.Elements<Paragraph>().ToList();
        Assert.Collection(commentParagraphs,
            firstParagraph => {
                Assert.Equal("First bold", string.Concat(firstParagraph.Descendants<Text>().Select(text => text.Text)));
                Assert.Contains(firstParagraph.Elements<Run>(), item => item.InnerText == "bold" && item.RunProperties?.Bold != null);
            },
            secondParagraph => {
                Assert.Equal("Second", string.Concat(secondParagraph.Descendants<Text>().Select(text => text.Text)));
                Assert.Contains(secondParagraph.Elements<Run>(), item => item.InnerText == "Second" && item.RunProperties?.Italic != null);
            });
    }

    [Fact]
    public void Word_Rtf_Bridge_Carries_Tab_Stops_And_Tab_Text() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph("Name");
        paragraph.AddTabStop(1440);
        paragraph.AddTabStop(2880, TabStopValues.Right, TabStopLeaderCharValues.Dot);
        paragraph.AddText("\tAmount");

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        using WordDocument roundTrip = rtf.LoadFromRtf();

        RtfParagraph rtfParagraph = Assert.Single(rtfDocument.Paragraphs);
        Assert.Equal("Name\tAmount", rtfParagraph.ToPlainText());
        Assert.Collection(rtfParagraph.TabStops,
            tabStop => {
                Assert.Equal(1440, tabStop.PositionTwips);
                Assert.Equal(RtfTabAlignment.Left, tabStop.Alignment);
                Assert.Equal(RtfTabLeader.None, tabStop.Leader);
            },
            tabStop => {
                Assert.Equal(2880, tabStop.PositionTwips);
                Assert.Equal(RtfTabAlignment.Right, tabStop.Alignment);
                Assert.Equal(RtfTabLeader.Dots, tabStop.Leader);
            });
        Assert.Contains(@"\tx1440", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\tldot\tqr\tx2880", rtf, StringComparison.Ordinal);
        Assert.Contains(@"Name\tab Amount", rtf, StringComparison.Ordinal);

        WordParagraph roundTripParagraph = Assert.Single(roundTrip.Paragraphs.GroupBy(item => item._paragraph).Select(group => group.First()));
        Assert.Equal(2, roundTripParagraph.TabStops.Count);
        Assert.Equal(1440, roundTripParagraph.TabStops[0].Position);
        Assert.Equal(TabStopValues.Right, roundTripParagraph.TabStops[1].Alignment);
        Assert.Equal(TabStopLeaderCharValues.Dot, roundTripParagraph.TabStops[1].Leader);
    }

    [Fact]
    public void Word_Rtf_Bridge_Carries_Explicit_Breaks() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph("Before");
        paragraph.AddBreak();
        paragraph.AddText("Line");
        paragraph.AddBreak(BreakValues.Page);
        paragraph.AddText("Page");
        paragraph.AddBreak(BreakValues.Column);
        paragraph.AddText("Column");

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        using WordDocument roundTrip = rtf.LoadFromRtf();

        RtfParagraph rtfParagraph = Assert.Single(rtfDocument.Paragraphs);
        Assert.Contains(rtfParagraph.Inlines, inline => inline is RtfBreak { Kind: RtfBreakKind.Line });
        Assert.Contains(rtfParagraph.Inlines, inline => inline is RtfBreak { Kind: RtfBreakKind.Page });
        Assert.Contains(rtfParagraph.Inlines, inline => inline is RtfBreak { Kind: RtfBreakKind.Column });
        Assert.Contains(@"Before\line Line\page Page\column Column", rtf, StringComparison.Ordinal);

        List<WordBreak> breaks = roundTrip.Breaks;
        Assert.Contains(breaks, item => item.BreakType == null);
        Assert.Contains(breaks, item => item.BreakType == BreakValues.Page);
        Assert.Contains(breaks, item => item.BreakType == BreakValues.Column);
    }

    [Fact]
    public void Rtf_Word_Bridge_Degrades_Soft_Breaks_To_Closest_Word_Breaks() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph("Before");
        paragraph.AddSoftLineBreak();
        paragraph.AddText("SoftLine");
        paragraph.AddSoftPageBreak();
        paragraph.AddText("SoftPage");

        using WordDocument word = document.ToWordDocument();

        List<WordBreak> breaks = word.Breaks;
        Assert.Contains(breaks, item => item.BreakType == null);
        Assert.Contains(breaks, item => item.BreakType == BreakValues.Page);
    }


}
