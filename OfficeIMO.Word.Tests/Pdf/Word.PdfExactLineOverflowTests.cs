using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(100D, false, false, false)]
    [InlineData(100D, true, false, false)]
    [InlineData(100D, false, true, false)]
    [InlineData(100D, true, true, false)]
    [InlineData(160D, false, false, false)]
    [InlineData(160D, true, false, false)]
    [InlineData(160D, false, true, false)]
    [InlineData(160D, true, true, false)]
    [InlineData(200D, false, false, false)]
    [InlineData(200D, true, false, false)]
    [InlineData(200D, false, true, false)]
    [InlineData(200D, true, true, false)]
    [InlineData(100D, false, false, true)]
    [InlineData(200D, true, true, true)]
    public void SaveAsPdf_ExactFirstLinesRetainRunningStoriesWhenTheyOverflow(double height, bool header, bool footer, bool prefix) {
        using WordDocument source = WordDocument.Create();
        source.Sections[0].Margins.HeaderDistance = source.Sections[0].Margins.FooterDistance = global::OfficeIMO.OfficeLength.Points(36D);
        if (header) SetStory(source.HeaderDefaultOrCreate);
        if (footer) SetStory(source.FooterDefaultOrCreate);
        if (prefix) {
            WordParagraph preceding = source.AddParagraph("P");
            preceding.FontSize = 12;
            preceding.LineSpacingPoints = 12;
            preceding.LineSpacingRule = WordLineSpacingRule.Exact;
            preceding.LineSpacingBefore = preceding.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
        }
        WordParagraph paragraph = source.AddParagraph();
        paragraph._paragraph.RemoveAllChildren<W.Run>();
        paragraph._paragraph.Append(TextLine("A", 12, true), TextLine("B", 12, false));
        paragraph.LineSpacingPoints = height;
        paragraph.LineSpacingRule = WordLineSpacingRule.Exact;
        paragraph.LineSpacingBefore = paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            PageSize = new OfficeIMO.Pdf.PdfPageSize(OfficeIMO.OfficeLength.Points(300),OfficeIMO.OfficeLength.Points(300)), Margins = PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(72))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(prefix ? 3 : 2, pdf.NumberOfPages);
        if (prefix) Assert.Equal("P", Assert.Single(pdf.GetPage(1).Letters).Value);
        for (int page = 1; page <= 2; page++) {
            var letter = Assert.Single(pdf.GetPage(page + (prefix ? 1 : 0)).Letters);
            Assert.Equal(page == 1 ? "A" : "B", letter.Value);
            // Native Word retains the header's 36 + 96 point exclusion on
            // initial and continuation pages, including ink in the bottom margin.
            Assert.Equal(300D - (header ? 132D : 72D) - height * .8D, letter.StartBaseLine.Y, 3);
        }

        static void SetStory(WordHeaderFooter story) {
            for (int index = 0; index < 3; index++) {
                WordParagraph blank = story.AddParagraph();
                blank.LineSpacingPoints = 32;
                blank.LineSpacingRule = WordLineSpacingRule.Exact;
                blank.LineSpacingBefore = blank.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
            }
        }
    }
}
