using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    public static IEnumerable<object[]> ParagraphBorderGroupFrameStartCases() {
        foreach (WordFileFormat format in new[] { WordFileFormat.Docx, WordFileFormat.Doc })
            foreach (WordCompatibilityMode mode in new[] { WordCompatibilityMode.Word2010, WordCompatibilityMode.Word2013 })
                foreach (var separator in new[] { (4U, 0U, 0U), (8U, 0U, 0U), (16U, 0U, 0U), (0U, 0U, 0U), (8U, 3U, 0U), (8U, 0U, 3U) })
                    yield return new object[] { format, mode, separator.Item1, separator.Item2, separator.Item3 };
    }

    [Theory]
    [MemberData(nameof(ParagraphBorderGroupFrameStartCases))]
    public void SaveAsPdf_ParagraphBorderGroupAtFrameStartUsesThePreviousSeparator(
        WordFileFormat format, WordCompatibilityMode mode, uint previousSize, uint previousSpace, uint currentSpace) {
        using WordDocument source = CreateParagraphBorderGroupFrameDocument(mode, 13, currentSpace);
        W.ParagraphBorders borders = source.Paragraphs[1]._paragraph.ParagraphProperties!.ParagraphBorders!;
        if (previousSize == 0) borders.BetweenBorder!.Remove();
        else { borders.BetweenBorder!.Size = previousSize; borders.BetweenBorder.Space = previousSpace; }
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(format)));
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, FontFamily = "Helvetica",
            PageSize = new OfficeIMO.Pdf.PdfPageSize(OfficeIMO.OfficeLength.Points(468),OfficeIMO.OfficeLength.Points(312)), Margins = OfficeIMO.Pdf.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(36))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        double before = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "Before").Letters[0].StartBaseLine.Y;
        var page = pdf.GetPage(2);
        double second = Assert.Single(page.GetWords(), word => word.Text == "Second01").Letters[0].StartBaseLine.Y;
        Assert.Equal(previousSize == 0 ? 0D : previousSpace + previousSize / 8D, before - second, 3);
        double last = Assert.Single(page.GetWords(), word => word.Text == "Second13").Letters[0].StartBaseLine.Y;
        Assert.DoesNotContain(page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle()),
            bounds => bounds.HasValue && bounds.Value.Width > 300D && bounds.Value.Height < .01D && bounds.Value.Bottom < last);
        var topStrokes = page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle())
            .Where(bounds => bounds.HasValue && bounds.Value.Width > 300D && bounds.Value.Height < .01D &&
                bounds.Value.Bottom > page.Height - 39D).ToArray();
        if (previousSize == 0) Assert.Empty(topStrokes);
        else {
            var stroke = Assert.Single(topStrokes)!.Value;
            Assert.Equal(page.Height - 36D - previousSize / 16D, stroke.Bottom, 3);
        }
    }

    [Theory]
    [InlineData(WordCompatibilityMode.Word2010)]
    [InlineData(WordCompatibilityMode.Word2013)]
    public void SaveAsPdf_ParagraphBorderGroupContinuationDoesNotRepeatTheSeparator(WordCompatibilityMode mode) {
        using WordDocument document = CreateParagraphBorderGroupFrameDocument(mode, 30, 3);
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, FontFamily = "Helvetica",
            PageSize = new OfficeIMO.Pdf.PdfPageSize(OfficeIMO.OfficeLength.Points(468),OfficeIMO.OfficeLength.Points(300)), Margins = OfficeIMO.Pdf.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(36))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var page = pdf.GetPage(2);
        Assert.StartsWith("First", page.GetWords().First().Text);
        double before = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "Before").Letters[0].StartBaseLine.Y;
        Assert.Equal(before, page.Letters[0].StartBaseLine.Y, 3);
        Assert.DoesNotContain(page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle()),
            bounds => bounds.HasValue && bounds.Value.Width > 300D && bounds.Value.Height < .01D &&
                bounds.Value.Bottom > page.Height - 39D);
    }

    [Theory]
    [InlineData(false, 24)]
    [InlineData(false, 48)]
    [InlineData(true, 0)]
    [InlineData(true, 24)]
    public void SaveAsPdf_JoinedParagraphSpacingAndBookmarksPreserveFrameSeparators(bool bookmark, int spacingAfter) {
        using WordDocument source = CreateParagraphBorderGroupFrameDocument(WordCompatibilityMode.Word2013, 13, 3);
        source.Paragraphs[1].LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(spacingAfter);
        if (bookmark) {
            W.Paragraph paragraph = source.Paragraphs.First(p => p.Text.StartsWith("Second", StringComparison.Ordinal))._paragraph;
            paragraph.InsertAt(new W.BookmarkStart { Id = "1", Name = "SecondParagraph" }, 1);
            paragraph.Append(new W.BookmarkEnd { Id = "1" });
        }
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, FontFamily = "Helvetica",
            PageSize = new OfficeIMO.Pdf.PdfPageSize(OfficeIMO.OfficeLength.Points(468),OfficeIMO.OfficeLength.Points(312)), Margins = OfficeIMO.Pdf.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(36))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(3, pdf.NumberOfPages);
        double before = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "Before").Letters[0].StartBaseLine.Y;
        var secondPage = pdf.GetPage(2);
        double second = Assert.Single(secondPage.GetWords(), word => word.Text == "Second01").Letters[0].StartBaseLine.Y;
        Assert.Equal(4D, before - second, 3);
        double firstLast = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "First13").Letters[0].StartBaseLine.Y;
        Assert.DoesNotContain(pdf.GetPage(1).Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle()),
            bounds => bounds.HasValue && bounds.Value.Width > 300D && bounds.Value.Height < .01D && bounds.Value.Bottom < firstLast);
        Assert.Single(secondPage.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle()),
            bounds => bounds.HasValue && bounds.Value.Width > 300D && bounds.Value.Height < .01D &&
                bounds.Value.Bottom > secondPage.Height - 39D);
    }

    private static WordDocument CreateParagraphBorderGroupFrameDocument(WordCompatibilityMode mode, int lines, uint space) {
        WordDocument document = WordDocument.Create();
        document.CompatibilitySettings.CompatibilityMode = mode;
        ConfigureAsymmetricBorderParagraph(document.AddParagraph("Before"));
        foreach (string prefix in new[] { "First", "Second", "Third" }) {
            WordParagraph paragraph = document.AddParagraph(string.Join("\n", Enumerable.Range(1, lines).Select(i => $"{prefix}{i:D2}")));
            ConfigureAsymmetricBorderParagraph(paragraph);
            ConfigureMatchingParagraphBorders(paragraph, "box");
            paragraph._paragraph.ParagraphProperties!.ParagraphBorders!.BetweenBorder =
                new W.BetweenBorder { Val = W.BorderValues.Single, Size = 8, Space = space, Color = "000000" };
        }
        ConfigureAsymmetricBorderParagraph(document.AddParagraph("After"));
        return document;
    }
}
