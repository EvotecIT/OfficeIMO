using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false, 1)]
    [InlineData(false, true, 1)]
    [InlineData(false, false, 4)]
    [InlineData(false, true, 4)]
    [InlineData(true, false, 1)]
    [InlineData(true, true, 1)]
    [InlineData(true, false, 4)]
    [InlineData(true, true, 4)]
    public void SaveAsPdf_MultipleSpacingFitsTheLastFontBoxInDocxAndDoc(bool nativeDoc, bool keep, int count) {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph();
        paragraph._paragraph.RemoveAllChildren<Run>();
        for (int index = 0; index < count; index++)
            paragraph._paragraph.Append(TextLine(((char)('A' + index)).ToString(), 10, index + 1 < count));
        paragraph.FontFamily = "Arial"; paragraph.FontSize = global::OfficeIMO.OfficeLength.Points(10);
        paragraph.LineSpacing = 480; paragraph.LineSpacingRule = WordLineSpacingRule.Auto;
        paragraph.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero; paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
        paragraph.AvoidWidowAndOrphanOverride = false;
        paragraph._paragraph.ParagraphProperties!.KeepLines = new KeepLines { Val = keep };
        paragraph._paragraph.ParagraphProperties.ParagraphMarkRunProperties = new ParagraphMarkRunProperties(
            new RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize { Val = "20" });
        using WordDocument document = WordDocument.Load(new MemoryStream(nativeDoc ? source.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc }) : source.ToBytes()));
        var options = new PdfCreateOptions { DefaultFont = PdfStandardFont.Helvetica };
        options.EmbedStandardFont(PdfStandardFont.Helvetica, CreateBaselineMetricFont(), "OfficeIMO-LineFit-Regular");
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(), PdfOptions = options,
            PageSize = new OfficeIMO.Pdf.PdfPageSize(OfficeIMO.OfficeLength.Points(200D),OfficeIMO.OfficeLength.Points((count == 1 ? 20D : 91D) + 40D)), Margins = PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(20D))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var page = Assert.Single(pdf.GetPages());
        var letters = page.Letters.Where(letter => "ABCD".Contains(letter.Value)).ToArray();
        Assert.Equal(count, letters.Length);
        for (int index = 1; index < letters.Length; index++)
            Assert.Equal(24D, letters[index - 1].StartBaseLine.Y - letters[index].StartBaseLine.Y, 3);
    }
}
