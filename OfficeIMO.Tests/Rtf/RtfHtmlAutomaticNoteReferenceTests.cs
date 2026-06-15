using System.Text;
using OfficeIMO.Html;
using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlAutomaticNoteReferenceTests {
    [Fact]
    public void RtfDocument_ToHtml_RoundTrips_Automatic_Note_Reference() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph("Value");
        var note = new RtfNote(RtfNoteKind.Footnote);
        RtfParagraph noteParagraph = note.AddParagraph("Footnote ");
        noteParagraph.AddText("text").SetItalic();
        paragraph.AddNoteReference(note);

        string html = document.ToHtml();

        Assert.Contains("data-officeimo-rtf-generated-text=\"note-reference\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-rtf-note=\"footnote\"", html, StringComparison.Ordinal);

        RtfDocument roundTrip = html.LoadFromHtml();
        RtfGeneratedText reference = Assert.IsType<RtfGeneratedText>(Assert.Single(roundTrip.Paragraphs).Inlines[1]);

        Assert.Equal(RtfGeneratedTextKind.NoteReference, reference.Kind);
        Assert.NotNull(reference.Note);
        Assert.Equal(RtfNoteKind.Footnote, reference.Note!.Kind);
        Assert.Equal("Footnote text", reference.Note.ToPlainText());
        Assert.Contains(reference.Note.Paragraphs[0].Runs, run => run.Text == "text" && run.Italic);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Automatic_Note_Reference_Metadata() {
        string content = Convert.ToBase64String(Encoding.UTF8.GetBytes("<p>Endnote <strong>text</strong></p>"));
        string html = "<p>Value<span data-officeimo-rtf-generated-text=\"note-reference\"></span><span data-officeimo-rtf-note=\"endnote\" data-officeimo-rtf-note-content=\"" + content + "\"></span></p>";

        RtfDocument document = html.LoadFromHtml();
        RtfGeneratedText reference = Assert.IsType<RtfGeneratedText>(Assert.Single(document.Paragraphs).Inlines[1]);

        Assert.Equal(RtfGeneratedTextKind.NoteReference, reference.Kind);
        Assert.NotNull(reference.Note);
        Assert.Equal(RtfNoteKind.Endnote, reference.Note!.Kind);
        Assert.Equal("Endnote text", reference.Note.ToPlainText());

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        Assert.Contains(@"\chftn {\endnote", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\b text\b0", rtf, StringComparison.Ordinal);
    }
}
