using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    public static IEnumerable<object[]> AsymmetricParagraphBorderCases() {
        foreach (WordFileFormat format in new[] { WordFileFormat.Docx, WordFileFormat.Doc }) {
            foreach (WordCompatibilityMode mode in new[] { WordCompatibilityMode.Word2010, WordCompatibilityMode.Word2013 }) {
                yield return new object[] { format, mode, "box-bottom", 1D, 4D };
                yield return new object[] { format, mode, "box-top", 4D, 1D };
                yield return new object[] { format, mode, "top-left", 4D, 0D };
                yield return new object[] { format, mode, "bottom-right", 0D, 4D };
            }
        }
    }

    [Theory]
    [MemberData(nameof(AsymmetricParagraphBorderCases))]
    public void SaveAsPdf_ParagraphBorderSpacingAdvancesOnlyItsOwnSide(
        WordFileFormat format, WordCompatibilityMode mode, string shape, double topAdvance, double bottomAdvance) {
        using WordDocument source = WordDocument.Create();
        source.CompatibilitySettings.CompatibilityMode = mode;
        WordParagraph before = source.AddParagraph("Before");
        WordParagraph middle = source.AddParagraph("Middle");
        WordParagraph after = source.AddParagraph("After");
        foreach (WordParagraph paragraph in new[] { before, middle, after }) {
            ConfigureAsymmetricBorderParagraph(paragraph);
        }
        bool box = shape.StartsWith("box-", StringComparison.Ordinal);
        middle.Borders.TopStyle = box || shape == "top-left" ? WordBorderStyle.Single : WordBorderStyle.Nil;
        middle.Borders.BottomStyle = box || shape == "bottom-right" ? WordBorderStyle.Single : WordBorderStyle.Nil;
        middle.Borders.LeftStyle = box || shape == "top-left" ? WordBorderStyle.Single : WordBorderStyle.Nil;
        middle.Borders.RightStyle = box || shape == "bottom-right" ? WordBorderStyle.Single : WordBorderStyle.Nil;
        middle.Borders.TopSize = middle.Borders.BottomSize = middle.Borders.LeftSize = middle.Borders.RightSize = 8;
        middle.Borders.TopSpace = shape == "box-top" || shape == "top-left" ? 3U : 0U;
        middle.Borders.BottomSpace = shape == "box-bottom" || shape == "bottom-right" ? 3U : 0U;
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(format)));
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false,
            PageSize = new OfficeIMO.Pdf.PdfPageSize(OfficeIMO.OfficeLength.Points(360),OfficeIMO.OfficeLength.Points(200)),
            Margins = OfficeIMO.Pdf.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(30)),
            FontFamily = "Helvetica"
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(1, pdf.NumberOfPages);
        var words = pdf.GetPage(1).GetWords().ToArray();
        double first = Assert.Single(words, word => word.Text == "Before").Letters[0].StartBaseLine.Y;
        double bordered = Assert.Single(words, word => word.Text == "Middle").Letters[0].StartBaseLine.Y;
        double last = Assert.Single(words, word => word.Text == "After").Letters[0].StartBaseLine.Y;
        Assert.Equal(16D + topAdvance, first - bordered, 3);
        Assert.Equal(16D + bottomAdvance, bordered - last, 3);
    }

    [Theory]
    [InlineData(WordFileFormat.Docx)]
    [InlineData(WordFileFormat.Doc)]
    public void SaveAsPdf_TopAndSideBorderDoesNotReserveMissingBottomBorderOnContinuation(WordFileFormat format) {
        using WordDocument source = WordDocument.Create();
        source.CompatibilitySettings.CompatibilityMode = WordCompatibilityMode.Word2010;
        ConfigureAsymmetricBorderParagraph(source.AddParagraph("Before"));
        WordParagraph paragraph = source.AddParagraph(string.Join("\n", Enumerable.Range(1, 30).Select(i => $"Marker{i:D3}")));
        ConfigureAsymmetricBorderParagraph(paragraph);
        paragraph.Borders.TopStyle = paragraph.Borders.LeftStyle = WordBorderStyle.Single;
        paragraph.Borders.TopSize = paragraph.Borders.LeftSize = 8;
        paragraph.Borders.TopSpace = 3;
        ConfigureAsymmetricBorderParagraph(source.AddParagraph("After"));
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(format)));
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false,
            PageSize = new OfficeIMO.Pdf.PdfPageSize(OfficeIMO.OfficeLength.Points(468),OfficeIMO.OfficeLength.Points(300)),
            Margins = OfficeIMO.Pdf.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(36)),
            FontFamily = "Helvetica"
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(3, pdf.NumberOfPages);
        for (int index = 1; index <= 30; index++) {
            int page = index <= 13 ? 1 : index <= 27 ? 2 : 3;
            Assert.Single(pdf.GetPage(page).GetWords(), word => word.Text == $"Marker{index:D3}");
        }
        Assert.Single(pdf.GetPage(3).GetWords(), word => word.Text == "After");
    }

    private static void ConfigureAsymmetricBorderParagraph(WordParagraph paragraph) {
        paragraph.FontFamily = "Arial";
        paragraph.FontSize = global::OfficeIMO.OfficeLength.Points(12);
        paragraph.LineSpacing = global::OfficeIMO.Word.WordLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(16));
        paragraph.LineSpacingBefore = paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
    }
}
