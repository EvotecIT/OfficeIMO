using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public sealed class RtfAuthoredNotePolicyRoundTripTests {
    [Theory]
    [InlineData("direct")]
    [InlineData("clone")]
    [InlineData("empty-append")]
    public void FirstFootnoteAddedAfterNoteFreeSaveRemainsAFootnote(string route) {
        RtfDocument original = RtfDocument.Create();
        original.AddParagraph("Body");
        RtfDocument source;
        switch (route) {
            case "clone":
                source = original.Clone();
                break;
            case "empty-append":
                source = RtfDocument.Create();
                source.AppendDocument(original).Report.RequireNoLoss();
                break;
            default:
                source = original;
                break;
        }
        Assert.Null(source.Settings.NoteMode);
        Assert.Empty(source.Notes);

        string saved = source.ToRtf();
        RtfDocument reopened = RtfDocument.Read(saved).Document;
        RtfNote added = Assert.Single(reopened.Paragraphs).AddFootnote("1", "Added after save").Note!;
        Assert.Equal(RtfNoteKind.Footnote, added.Kind);
        using WordDocument word = reopened.ToWordDocument();
        using var stream = new MemoryStream();
        word.Save(stream);
        stream.Position = 0;
        using WordprocessingDocument persisted = WordprocessingDocument.Open(stream, false);
        MainDocumentPart main = persisted.MainDocumentPart!;
        Assert.Single(main.Document!.Descendants<FootnoteReference>());
        Assert.Empty(main.Document.Descendants<EndnoteReference>());
        Footnote footnote = Assert.Single(main.FootnotesPart!.Footnotes!.Elements<Footnote>(), note => note.Id!.Value > 0);
        Assert.Equal("Added after save", Assert.Single(footnote.Elements<Paragraph>()).InnerText);
        Assert.Equal(RtfNoteKind.Footnote, added.Kind);
        Assert.Null(source.Settings.NoteMode);
        Assert.Empty(source.Notes);
        Assert.Equal("Body", Assert.Single(source.Paragraphs).ToPlainText());
    }
}
