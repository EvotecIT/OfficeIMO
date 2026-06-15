using System.Text;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlNoteTests {
    [Fact]
    public void RtfDocument_ToHtml_Renders_Footnote_Metadata_And_RoundTrips() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph("Value");
        RtfRun reference = paragraph.AddText("1").SetSuperscript();
        var note = new RtfNote(RtfNoteKind.Footnote);
        RtfParagraph noteParagraph = note.AddParagraph();
        noteParagraph.AddText("Footnote ");
        noteParagraph.AddText("text").SetItalic();
        reference.SetNote(note);

        string html = document.ToHtml();

        Assert.Contains("data-officeimo-rtf-note=\"footnote\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-rtf-note-content=\"", html, StringComparison.Ordinal);

        RtfDocument roundTrip = html.ToRtfDocumentFromHtml();
        RtfRun roundTripReference = Assert.Single(roundTrip.Paragraphs).Runs[1];
        Assert.NotNull(roundTripReference.Note);
        Assert.Equal(RtfNoteKind.Footnote, roundTripReference.Note!.Kind);
        Assert.Equal("Footnote text", roundTripReference.Note.ToPlainText());
        Assert.Contains(roundTripReference.Note.Paragraphs[0].Runs, run => run.Text == "text" && run.Italic);

        string rtf = roundTrip.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        Assert.Contains(@"{\footnote", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\i text\i0", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Endnote_Metadata() {
        string content = Convert.ToBase64String(Encoding.UTF8.GetBytes("<p>Endnote <strong>text</strong></p>"));
        string html = "<p>Value<sup>i</sup><span data-officeimo-rtf-note=\"endnote\" data-officeimo-rtf-note-content=\"" + content + "\"></span></p>";

        RtfDocument document = html.ToRtfDocumentFromHtml();
        RtfRun reference = Assert.Single(document.Paragraphs).Runs[1];

        Assert.NotNull(reference.Note);
        Assert.Equal(RtfNoteKind.Endnote, reference.Note!.Kind);
        Assert.Equal("Endnote text", reference.Note.ToPlainText());
        Assert.Contains(reference.Note.Paragraphs[0].Runs, run => run.Text == "text" && run.Bold);

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        Assert.Contains(@"{\endnote", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\b text\b0", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Annotation_Metadata() {
        string content = Convert.ToBase64String(Encoding.UTF8.GetBytes("<p>Review note</p>"));
        const string created = "2026-01-02T03:04:05.0000000";
        string html = "<p>Flag<span data-officeimo-rtf-note=\"annotation\" data-officeimo-rtf-note-content=\"" + content + "\" data-officeimo-rtf-note-id=\"c1\" data-officeimo-rtf-note-author=\"Alice\" data-officeimo-rtf-note-created=\"" + created + "\"></span></p>";

        RtfDocument document = html.ToRtfDocumentFromHtml();
        RtfNote note = Assert.Single(document.Paragraphs).Runs[0].Note!;

        Assert.Equal(RtfNoteKind.Annotation, note.Kind);
        Assert.Equal("c1", note.Id);
        Assert.Equal("Alice", note.Author);
        Assert.Equal(new DateTime(2026, 1, 2, 3, 4, 5), note.Created);
        Assert.Equal("Review note", note.ToPlainText());

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        Assert.Contains(@"{\annotation", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\*\atnid c1}", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\*\atnauthor Alice}", rtf, StringComparison.Ordinal);
    }
}
