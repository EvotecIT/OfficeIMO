using OfficeIMO.Pdf;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfPdfAutomaticNoteReferenceTests {
    [Fact]
    public void RtfDocument_ToPdfDocument_Renders_Automatic_Note_Bodies() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph("Body");
        var note = new RtfNote(RtfNoteKind.Footnote);
        note.AddParagraph("Footnote body");
        paragraph.AddNoteReference(note, "1");

        string text = PdfReadDocument.Open(document.ToPdf()).ExtractText();

        Assert.Contains("Body1", text, StringComparison.Ordinal);
        Assert.Contains("Footnote 1:", text, StringComparison.Ordinal);
        Assert.Contains("Footnote body", text, StringComparison.Ordinal);
    }
}
