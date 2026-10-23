using DocumentFormat.OpenXml;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void SaveAsPdf_LoadedNoteParagraphBookmarkResolvesAtItsActualTarget(bool endnote, bool numbered) {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        OpenXmlCompositeElement note = AddRichNoteFixture(source, endnote);
        WordParagraph target = numbered
            ? source.AddList(WordListStyle.Numbered).AddItem("NOTETARGET")
            : source.AddParagraph("NOTETARGET");
        target._paragraph.Remove();
        note.Append(target._paragraph);
        target.AddBookmark("LoadedNoteTarget");
        source.AddParagraph(" ")._paragraph.Append(new W.Hyperlink(RichNoteRun("OPENBOOKMARK")) { Anchor = "LoadedNoteTarget" });
        Assert.Empty(source.ValidateDocument());

        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        var main = document._wordprocessingDocument.MainDocumentPart!;
        OpenXmlCompositeElement loadedNotes = endnote ? main.EndnotesPart!.Endnotes! : main.FootnotesPart!.Footnotes!;
        Assert.Single(loadedNotes.Descendants<W.BookmarkStart>(), bookmark => bookmark.Name?.Value == "LoadedNoteTarget");
        string original = main.Document.OuterXml + loadedNotes.OuterXml;
        PdfDocumentConversionResult result = document.ToPdfDocumentResult(FootnoteDiagnosticOptions());
        for (int serialization = 0; serialization < 2; serialization++) {
            byte[] bytes = result.ToBytes();
            using var pdf = PdfPigDocument.Open(bytes);
            PdfDocumentInfo navigation = PdfInspector.Inspect(bytes);
            Assert.Single(navigation.LinkAnnotations, link => link.DestinationName == "LoadedNoteTarget");
            PdfNamedDestination destination = Assert.Single(navigation.NamedDestinations, item => item.Name == "LoadedNoteTarget");
            var page = pdf.GetPages().Single(item => item.Text.Contains("NOTETARGET", StringComparison.Ordinal));
            Assert.Equal(page.Number, Assert.IsType<int>(destination.PageNumber));
            var letters = FindRichNoteToken(page.Letters.ToArray(), "NOTETARGET");
            double top = Assert.IsType<double>(destination.DestinationTop);
            Assert.InRange(top - letters[0].StartBaseLine.Y, 0D, 40D);
            if (numbered) Assert.Contains("1.", page.Text, StringComparison.Ordinal);
            Assert.DoesNotContain(result.Warnings, warning => warning.Code is
                "NativeFootnotePlacementApproximated" or "NativeEndnoteContinuationApproximated" or "NativeEndnoteContentApproximated");
            Assert.Equal(original, main.Document.OuterXml + loadedNotes.OuterXml);
        }
        Assert.Empty(document.ValidateDocument());
    }
}
