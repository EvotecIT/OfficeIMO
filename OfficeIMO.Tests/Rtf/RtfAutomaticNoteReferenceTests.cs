using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfAutomaticNoteReferenceTests {
    [Fact]
    public void Read_And_Write_Automatic_Footnote_Reference() {
        const string rtf = @"{\rtf1\ansi\pard Body\chftn{\footnote\pard Footnote {\i text}\par} after\par}";

        RtfReadResult result = RtfDocument.Read(rtf);
        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);

        Assert.Equal("Body after", paragraph.ToPlainText());
        Assert.Collection(paragraph.Inlines,
            inline => Assert.Equal("Body", Assert.IsType<RtfRun>(inline).Text),
            inline => {
                RtfGeneratedText reference = Assert.IsType<RtfGeneratedText>(inline);
                Assert.Equal(RtfGeneratedTextKind.NoteReference, reference.Kind);
                Assert.NotNull(reference.Note);
                Assert.Equal(RtfNoteKind.Footnote, reference.Note!.Kind);
                Assert.Equal("Footnote text", reference.Note.ToPlainText());
                Assert.Contains(reference.Note.Paragraphs[0].Runs, run => run.Text == "text" && run.Italic);
            },
            inline => Assert.Equal(" after", Assert.IsType<RtfRun>(inline).Text));
        Assert.Same(Assert.IsType<RtfGeneratedText>(paragraph.Inlines[1]).Note, Assert.Single(result.Document.Notes));

        string written = result.Document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });

        Assert.Contains(@"\chftn {\footnote", written, StringComparison.Ordinal);
        Assert.Contains(@"\i text\i0", written, StringComparison.Ordinal);
    }

    [Fact]
    public void AddNoteReference_Emits_Automatic_Endnote_Reference() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph("Body");
        var note = new RtfNote(RtfNoteKind.Endnote);
        note.AddParagraph("Endnote body");
        paragraph.AddNoteReference(note);

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfDocument roundTrip = RtfDocument.Read(rtf).Document;
        RtfGeneratedText reference = Assert.IsType<RtfGeneratedText>(Assert.Single(roundTrip.Paragraphs).Inlines[1]);

        Assert.Contains(@"\chftn {\endnote", rtf, StringComparison.Ordinal);
        Assert.Equal(RtfGeneratedTextKind.NoteReference, reference.Kind);
        Assert.NotNull(reference.Note);
        Assert.Equal(RtfNoteKind.Endnote, reference.Note!.Kind);
        Assert.Equal("Endnote body", reference.Note.ToPlainText());
    }
}
