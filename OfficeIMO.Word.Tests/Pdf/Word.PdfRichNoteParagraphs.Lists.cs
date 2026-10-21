using DocumentFormat.OpenXml;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_DirectNoteListRetainsCanonicalMarkersAndContinuationIndent(bool endnote) {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        OpenXmlCompositeElement note = AddRichNoteFixture(source, endnote);
        note.RemoveAllChildren();
        WordList list = source.AddList(WordListStyle.Numbered);
        WordParagraph first = list.AddItem("FIRSTITEM");
        WordParagraph second = list.AddItem("SECONDITEM");
        WordParagraph third = list.AddItem("THIRDITEM");
        foreach (WordParagraph item in new[] { first, second, third }) {
            item._paragraph.Remove();
            note.Append(item._paragraph);
        }
        first._paragraph.InsertAfter(NoteBodyMarker(endnote), first._paragraph.ParagraphProperties!);
        Assert.Empty(source.ValidateDocument());
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        var main = document._wordprocessingDocument.MainDocumentPart!;
        string original = (endnote ? (OpenXmlElement)main.EndnotesPart!.Endnotes! : main.FootnotesPart!.Footnotes!).OuterXml;
        var result = document.ToPdfDocumentResult(FootnoteDiagnosticOptions());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("1.", text, StringComparison.Ordinal);
        Assert.Contains("2.", text, StringComparison.Ordinal);
        Assert.Contains("3.", text, StringComparison.Ordinal);
        var letters = pdf.GetPages().SelectMany(page => page.Letters).ToArray();
        FindRichNoteToken(letters, "FIRSTITEM");
        var middle = FindRichNoteToken(letters, "SECONDITEM");
        var last = FindRichNoteToken(letters, "THIRDITEM");
        Assert.Equal(middle[0].StartBaseLine.X, last[0].StartBaseLine.X, 3);
        Assert.True(middle[0].StartBaseLine.Y > last[0].StartBaseLine.Y);
        Assert.DoesNotContain(result.Warnings, warning => warning.Code is
            "NativeFootnotePlacementApproximated" or "NativeEndnoteContentApproximated");
        Assert.Equal(original, (endnote ? (OpenXmlElement)main.EndnotesPart.Endnotes! : main.FootnotesPart!.Footnotes!).OuterXml);
    }

    [Fact]
    public void SaveAsPdf_HeadingStyleInsideAnEndnoteRetainsSourceTypographyAsNoteParagraphFlow() {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        OpenXmlCompositeElement note = AddRichNoteFixture(source, endnote: true);
        source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(new W.Style(
            new W.StyleName { Val = "Note heading" }, new W.StyleParagraphProperties(new W.OutlineLevel { Val = 0 }),
            new W.StyleRunProperties(new W.Bold(), new W.Color { Val = "803040" }, new W.FontSize { Val = "32" })) {
                Type = W.StyleValues.Paragraph, StyleId = "NoteHeading" });
        note.RemoveAllChildren();
        note.Append(new W.Paragraph(new W.ParagraphProperties(new W.ParagraphStyleId { Val = "NoteHeading" }),
            NoteBodyMarker(endnote: true), RichNoteRun("HEADINGNOTE")), new W.Paragraph(RichNoteRun("FOLLOWINGNOTE")));
        Assert.Empty(source.ValidateDocument());
        string original = note.OuterXml;
        using var pdf = PdfPigDocument.Open(source.ToPdfBytes(FootnoteDiagnosticOptions()));
        var letters = pdf.GetPages().SelectMany(page => page.Letters).ToArray();
        var heading = FindRichNoteToken(letters, "HEADINGNOTE");
        var following = FindRichNoteToken(letters, "FOLLOWINGNOTE");
        Assert.All(heading, letter => {
            Assert.Contains("Bold", letter.FontName, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(16D, letter.PointSize, 3);
            Assert.Equal(128D / 255D, letter.Color.ToRGBValues().Item1, 3);
        });
        Assert.True(heading[0].StartBaseLine.Y > following[0].StartBaseLine.Y);
        Assert.True(heading[0].PointSize > following[0].PointSize);
        Assert.Equal(original, note.OuterXml);
    }
}
