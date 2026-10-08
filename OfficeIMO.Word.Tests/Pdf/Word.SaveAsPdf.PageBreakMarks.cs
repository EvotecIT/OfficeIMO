using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordCompatibilityMode.Word2003, false)]
    [InlineData(WordCompatibilityMode.Word2003, true)]
    [InlineData(WordCompatibilityMode.Word2013, false)]
    [InlineData(WordCompatibilityMode.Word2013, true)]
    public void SaveAsPdf_PageBreakMarkAddsOnlyTheLegacyTrailingMark(WordCompatibilityMode mode, bool precedingText) {
        using WordDocument disabled = CreatePageBreakMarkDocument(mode, false, precedingText);
        using WordDocument enabled = CreatePageBreakMarkDocument(mode, true, precedingText);
        using PdfPigDocument first = PdfPigDocument.Open(disabled.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        using PdfPigDocument second = PdfPigDocument.Open(enabled.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        Assert.Equal(2, first.NumberOfPages);
        Assert.Equal(2, second.NumberOfPages);
        Assert.Equal(first.GetPage(1).Text, second.GetPage(1).Text);
        Assert.Equal(first.GetPage(2).Text, second.GetPage(2).Text);
        double delta = first.GetPage(2).Letters[0].StartBaseLine.Y - second.GetPage(2).Letters[0].StartBaseLine.Y;
        Assert.InRange(delta, mode == WordCompatibilityMode.Word2013 ? -0.01 : 21.99,
            mode == WordCompatibilityMode.Word2013 ? 0.01 : 22.01);
    }

    [Fact]
    public void SaveAsPdf_PageBreakMarkNativeImportAndDocxProjectionKeepTheSameLayout() {
        using WordDocument source = CreatePageBreakMarkDocument(WordCompatibilityMode.Word2003, true, true);
        using WordDocument native = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        using WordDocument projected = WordDocument.Load(new MemoryStream(native.ToBytes(WordFileFormat.Docx)));
        Assert.True(native.CompatibilitySettings.SplitPageBreakAndParagraphMark);
        Assert.Equal(WordCompatibilityMode.Word2003, projected.CompatibilitySettings.CompatibilityMode);
        using PdfPigDocument first = PdfPigDocument.Open(native.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        using PdfPigDocument second = PdfPigDocument.Open(projected.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        Assert.Equal(2, first.NumberOfPages);
        Assert.Equal(2, second.NumberOfPages);
        for (int page = 1; page <= 2; page++) {
            Assert.Equal(first.GetPage(page).Text, second.GetPage(page).Text);
            Assert.Equal(first.GetPage(page).Letters.Select(letter => letter.StartBaseLine),
                second.GetPage(page).Letters.Select(letter => letter.StartBaseLine));
        }
    }

    [Fact]
    public void SaveAsPdf_PageBreakMarkDoesNotDuplicateAContentContinuation() {
        using WordDocument disabled = CreatePageBreakMarkDocument(WordCompatibilityMode.Word2003, false, true, true);
        using WordDocument enabled = CreatePageBreakMarkDocument(WordCompatibilityMode.Word2003, true, true, true);
        using PdfPigDocument first = PdfPigDocument.Open(disabled.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        using PdfPigDocument second = PdfPigDocument.Open(enabled.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        Assert.Equal(2, first.NumberOfPages);
        Assert.Equal(2, second.NumberOfPages);
        Assert.Contains("AFTER", second.GetPage(2).Text);
        Assert.Equal(first.GetPage(2).Text, second.GetPage(2).Text);
        Assert.Equal(first.GetPage(2).Letters.Select(letter => letter.StartBaseLine),
            second.GetPage(2).Letters.Select(letter => letter.StartBaseLine));
    }

    private static WordDocument CreatePageBreakMarkDocument(WordCompatibilityMode mode, bool enabled,
        bool precedingText, bool followingText = false) {
        WordDocument document = WordDocument.Create();
        document.CompatibilitySettings.CompatibilityMode = mode;
        document.CompatibilitySettings.SplitPageBreakAndParagraphMark = enabled;
        document.AddParagraph("FIRST");
        WordParagraph paragraph = document.AddParagraph();
        paragraph._paragraph.ParagraphProperties = new ParagraphProperties(
            new SpacingBetweenLines { Before = "0", After = "80", Line = "360", LineRule = LineSpacingRuleValues.Exact },
            new ParagraphMarkRunProperties(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize { Val = "24" }));
        paragraph._paragraph.Append(new Run(new RunProperties(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize { Val = "24" }),
            new Text(precedingText ? "BEFORE" : ""), new Break { Type = BreakValues.Page }));
        if (followingText) paragraph._paragraph.Append(new Run(new Text("AFTER")));
        document.AddParagraph("TARGET");
        return document;
    }
}
