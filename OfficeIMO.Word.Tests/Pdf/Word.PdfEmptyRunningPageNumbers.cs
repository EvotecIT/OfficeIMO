using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(0, false)]
    [InlineData(1, false)]
    [InlineData(2, false)]
    [InlineData(0, true)]
    [InlineData(1, true)]
    [InlineData(2, true)]
    public void SaveAsPdf_EmptyRunningFooterPreservesGeneratedPageNumberPresentation(int emptyVariant, bool typography) {
        using WordDocument document = CreateEmptyRunningStoryDocument("absent", footer: true);
        WordHeaderFooter[] stories = { document.FooterFirstOrCreate, document.FooterEvenOrCreate, document.FooterDefaultOrCreate };
        foreach (var child in stories[emptyVariant].ChildElements.ToArray()) child.Remove();
        for (int index = 0; index < 3; index++) {
            if (index != emptyVariant) stories[index].AddParagraph("STORY" + index);
            WordParagraph body = document.AddParagraph("A");
            body.LineSpacing = 240; body.LineSpacingRule = WordLineSpacingRule.Exact;
            body.LineSpacingBeforePoints = body.LineSpacingAfterPoints = 0;
            if (index < 2) body.AddBreak(WordBreakType.Page);
        }
        WordToPdfOptions options = EmptyRunningOptions();
        options.IncludePageNumbers = true;
        options.PageNumberFormat = "NUMBER {page}/{pages}";
        options.PdfOptions = new PdfOptions { FooterOffsetY = typography ? 18D : 24D };
        if (typography) {
            options.PdfOptions.FooterFont = PdfStandardFont.Courier;
            options.PdfOptions.FooterFontSize = 24D;
            options.PdfOptions.FooterAlign = PdfAlign.Left;
            options.PdfOptions.FooterTextColor = PdfColor.FromRgb(255, 0, 0);
        }
        using var withImplicit = PdfPigDocument.Open(document.ToPdfBytes(options));
        W.HeaderFooterValues kind = emptyVariant == 0 ? W.HeaderFooterValues.First
            : emptyVariant == 1 ? W.HeaderFooterValues.Even : W.HeaderFooterValues.Default;
        foreach (var reference in document.Sections[0]._sectionProperties.Elements<W.FooterReference>()
            .Where(reference => reference.Type?.Value == kind).ToArray()) reference.Remove();
        using WordDocument withoutDocument = WordDocument.Load(new MemoryStream(document.ToBytes()));
        using var withoutImplicit = PdfPigDocument.Open(withoutDocument.ToPdfBytes(options));
        Assert.Equal(3, withImplicit.NumberOfPages);
        Assert.Equal(3, withoutImplicit.NumberOfPages);
        for (int index = 0; index < 3; index++) {
            // Removing an empty source part must not restyle or reposition a
            // generated number, nor alter either nonempty sibling variant.
            Assert.Equal(withoutImplicit.GetPage(index + 1).Letters.Select(letter =>
                    (letter.Value, letter.StartBaseLine, letter.FontSize, letter.FontName, letter.Color.ToRGBValues())),
                withImplicit.GetPage(index + 1).Letters.Select(letter =>
                    (letter.Value, letter.StartBaseLine, letter.FontSize, letter.FontName, letter.Color.ToRGBValues())));
            if (index == emptyVariant && emptyVariant < 2)
                Assert.DoesNotContain("NUMBER", withImplicit.GetPage(index + 1).Text);
            else Assert.Contains($"NUMBER {index + 1}/3", withImplicit.GetPage(index + 1).Text);
        }
    }
}
