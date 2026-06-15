using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class WordRtfConverterTests {
    [Fact]
    public void Rtf_To_Word_Bridge_Carries_Automatic_Footnote_Reference() {
        RtfDocument rtfDocument = RtfDocument.Create();
        RtfParagraph paragraph = rtfDocument.AddParagraph("Body");
        var note = new RtfNote(RtfNoteKind.Footnote);
        note.AddParagraph("Footnote body");
        paragraph.AddNoteReference(note);

        using WordDocument word = rtfDocument.ToWordDocument();

        WordFootNote footNote = Assert.Single(word.FootNotes);
        Assert.NotNull(footNote.ParentParagraph);
        Assert.Contains(word.Paragraphs, item => item.Text == "Body");
        Assert.Contains(footNote.Paragraphs!, item => item.Text == "Footnote body");
        Assert.Single(word._document.Body!.Descendants<FootnoteReference>());
    }
}
