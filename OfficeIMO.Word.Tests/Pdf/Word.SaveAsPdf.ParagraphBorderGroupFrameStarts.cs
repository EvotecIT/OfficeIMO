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
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, FontFamily = "Helvetica",
            PageSize = new OfficeIMO.Pdf.PageSize(468, 312), Margins = OfficeIMO.Pdf.PageMargins.Uniform(36)
        }));
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
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, FontFamily = "Helvetica",
            PageSize = new OfficeIMO.Pdf.PageSize(468, 300), Margins = OfficeIMO.Pdf.PageMargins.Uniform(36)
        }));
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
        source.Paragraphs[1].LineSpacingAfterPoints = spacingAfter;
        if (bookmark) {
            W.Paragraph paragraph = source.Paragraphs.First(p => p.Text.StartsWith("Second", StringComparison.Ordinal))._paragraph;
            paragraph.InsertAt(new W.BookmarkStart { Id = "1", Name = "SecondParagraph" }, 1);
            paragraph.Append(new W.BookmarkEnd { Id = "1" });
        }
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, FontFamily = "Helvetica",
            PageSize = new OfficeIMO.Pdf.PageSize(468, 312), Margins = OfficeIMO.Pdf.PageMargins.Uniform(36)
        }));
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

    [Fact]
    public void SaveAsPdf_FinalJoinedParagraphReservesItsAutomaticLineAdvance() {
        using WordDocument source = WordDocument.Create();
        source.CompatibilitySettings.CompatibilityMode = WordCompatibilityMode.Word2013;
        WordParagraph[] paragraphs = new[] { "Lead", "A", "B" }.Select(source.AddParagraph).ToArray();
        foreach (WordParagraph paragraph in paragraphs) {
            ConfigureAsymmetricBorderParagraph(paragraph);
            paragraph.FontSize = 10;
            paragraph.AvoidWidowAndOrphanOverride = false;
            paragraph.KeepLinesTogetherOverride = false;
            paragraph._paragraph.ParagraphProperties!.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties(
                new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new W.FontSize { Val = "20" });
        }
        paragraphs[0].LineSpacingPoints = 30D;
        foreach (WordParagraph paragraph in paragraphs.Skip(1)) {
            paragraph.LineSpacing = 480;
            paragraph.LineSpacingRule = WordLineSpacingRule.Auto;
            ConfigureMatchingParagraphBorders(paragraph, "box");
            W.ParagraphBorders borders = paragraph._paragraph.ParagraphProperties!.ParagraphBorders!;
            foreach (W.BorderType border in borders.ChildElements.OfType<W.BorderType>()) border.Color = "000000";
            borders.BetweenBorder = new W.BetweenBorder {
                Val = W.BorderValues.Single, Size = 8, Space = 0, Color = "000000"
            };
        }
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        var pdfOptions = new OfficeIMO.Pdf.PdfOptions { DefaultFont = OfficeIMO.Pdf.PdfStandardFont.Helvetica };
        pdfOptions.EmbedStandardFont(OfficeIMO.Pdf.PdfStandardFont.Helvetica,
            CreateBaselineMetricFont(), "OfficeIMO-LineFit-Regular");
        using var output = new MemoryStream();
        document.SaveAsPdf(output, new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic(),
            PdfOptions = pdfOptions, PageSize = new OfficeIMO.Pdf.PageSize(200D, 120D),
            Margins = OfficeIMO.Pdf.PageMargins.Uniform(20D)
        });
        using var pdf = PdfPigDocument.Open(output.ToArray());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "Lead");
        Assert.Single(pdf.GetPage(1).Letters, letter => letter.Value == "A");
        Assert.DoesNotContain("B", pdf.GetPage(1).Text, StringComparison.Ordinal);
        var page = pdf.GetPage(2);
        Assert.Single(page.Letters, letter => letter.Value == "B");
        double[] horizontal = page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle())
            .Where(bounds => bounds.HasValue && bounds.Value.Width > 100D && bounds.Value.Height < .01D)
            .Select(bounds => bounds!.Value.Bottom).OrderBy(value => value).ToArray();
        Assert.Equal(2, horizontal.Length);
        // The final paragraph keeps its 24-point automatic advance and four-point
        // outer padding; the previous one-point separator starts the new frame.
        Assert.Equal(71.5D, horizontal[0], 3);
        Assert.Equal(99.5D, horizontal[1], 3);
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
