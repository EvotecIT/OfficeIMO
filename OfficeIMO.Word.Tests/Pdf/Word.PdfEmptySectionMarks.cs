using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, "pure")]
    [InlineData(false, "styled")]
    [InlineData(false, "hidden")]
    [InlineData(false, "mark-size")]
    [InlineData(true, "pure")]
    [InlineData(true, "styled")]
    public void SaveAsPdf_EmptySectionRetainsItsOwnParagraphLine(bool precedingText, string mode) {
        using WordDocument source = CreateDocument();
        if (precedingText) AddLine(source, "P");
        W.Body body = source._wordprocessingDocument.MainDocumentPart!.Document.Body!;
        W.SectionProperties final = body.Elements<W.SectionProperties>().Single();
        var properties = new W.ParagraphProperties((W.SectionProperties)final.CloneNode(true));
        ConfigureMark(properties);
        body.InsertBefore(new W.Paragraph(properties), final);
        final.RemoveAllChildren<W.SectionType>();
        final.AddChild(new W.SectionType { Val = W.SectionMarkValues.Continuous }, true);
        AddLine(source, "A").AddBreak(WordBreakType.Page);
        AddLine(source, "B");
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        string before = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;

        using WordDocument reference = CreateDocument();
        if (precedingText) AddLine(reference, "P");
        // Native Word gives an otherwise empty section its paragraph line.
        // A terminator after real text adds no separate line.
        else if (mode != "hidden") {
            WordParagraph blank = reference.AddParagraph();
            ConfigureMark(blank._paragraph.ParagraphProperties ??= new W.ParagraphProperties());
        }
        AddLine(reference, "A").AddBreak(WordBreakType.Page);
        AddLine(reference, "B");
        var options = new WordToPdfOptions { ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() };
        using var actual = PdfPigDocument.Open(document.ToPdfBytes(options));
        using var expected = PdfPigDocument.Open(reference.ToPdfBytes(options));
        Assert.Equal(2, actual.NumberOfPages);
        for (int page = 1; page <= 2; page++) Assert.Equal(
            expected.GetPage(page).Letters.Select(letter => (letter.Value, letter.StartBaseLine)),
            actual.GetPage(page).Letters.Select(letter => (letter.Value, letter.StartBaseLine)));
        Assert.Equal(before, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);

        void ConfigureMark(W.ParagraphProperties properties) {
            if (mode == "styled") properties.ParagraphStyleId = new W.ParagraphStyleId { Val = "Normal" };
            if (mode == "hidden") properties.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties(new W.Vanish());
            if (mode == "mark-size") properties.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties(new W.FontSize { Val = "48" });
        }
        static WordDocument CreateDocument() {
            var result = WordDocument.Create();
            result.CompatibilitySettings.CompatibilityMode = WordCompatibilityMode.Word2013;
            result.Settings.FontFamily = "Arial";
            result.Settings.FontSize = 12;
            return result;
        }
        static WordParagraph AddLine(WordDocument target, string text) {
            WordParagraph paragraph = target.AddParagraph(text);
            paragraph.LineSpacingPoints = 12;
            paragraph.LineSpacingRule = WordLineSpacingRule.Exact;
            paragraph.LineSpacingBeforePoints = paragraph.LineSpacingAfterPoints = 0;
            return paragraph;
        }
    }
}
