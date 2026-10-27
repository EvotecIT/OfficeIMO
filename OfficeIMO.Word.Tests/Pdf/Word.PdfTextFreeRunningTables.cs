using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void SaveAsPdf_TextFreeRunningTableRetainsBordersAndDrawsImageOnce(bool footer, bool image) {
        using WordDocument document = WordDocument.Create();
        WordHeaderFooter story = footer ? document.FooterDefaultOrCreate : document.HeaderDefaultOrCreate;
        WordTable table = story.AddTable(1, 2, WordTableStyle.TableGrid);
        if (image) table.Rows[0].Cells[0].Paragraphs[0].AddImage(Path.Combine(_directoryWithImages, "EvotecLogo.png"), 32, 32);
        document.AddParagraph("BODY");
        string before = string.Concat(story.ChildElements.Select(element => element.OuterXml));
        byte[] bytes = document.ToPdf(EmptyRunningOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Single(pdf.GetPages());
        Assert.Contains("BODY", pdf.GetPage(1).Text);
        Assert.Equal(image ? 1 : 0, pdf.GetPage(1).GetImages().Count());
        Assert.Contains(" RG", PdfOperatorSearchText.From(bytes));
        Assert.Equal(before, string.Concat(story.ChildElements.Select(element => element.OuterXml)));
    }

    [Theory]
    [InlineData("default", false)]
    [InlineData("default", true)]
    [InlineData("first", false)]
    [InlineData("first", true)]
    [InlineData("even", false)]
    [InlineData("even", true)]
    public void SaveAsPdf_TextFreeRunningTablePreservesGeneratedNumberSelection(string variant, bool image) {
        using WordDocument document = WordDocument.Create();
        WordHeaderFooter story = variant == "first" ? document.FooterFirstOrCreate
            : variant == "even" ? document.FooterEvenOrCreate : document.FooterDefaultOrCreate;
        // Isolate footer variants from header peers created by the fluent API.
        document.Sections[0]._sectionProperties.RemoveAllChildren<W.HeaderReference>();
        document.Sections[0].DifferentFirstPage = variant == "first";
        document.Sections[0].DifferentOddAndEvenPages = variant == "even";
        WordTable table = story.AddTable(1, 2, WordTableStyle.TableGrid);
        if (image) table.Rows[0].Cells[0].Paragraphs[0].AddImage(Path.Combine(_directoryWithImages, "EvotecLogo.png"), 32, 32);
        SetInheritedRunningControlStyle(document.AddParagraph("BODY"));
        WordParagraph second = document.AddParagraph("SECOND");
        SetInheritedRunningControlStyle(second);
        second._paragraph.ParagraphProperties!.AddChild(new W.PageBreakBefore(), true);
        string before = string.Concat(story.ChildElements.Select(element => element.OuterXml));
        WordPdfExportOptions options = EmptyRunningOptions();
        options.IncludePageNumbers = true;
        options.PageNumberFormat = "NUMBER {page}/{pages}";
        using var pdf = PdfPigDocument.Open(document.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(2, pdf.NumberOfPages);
        for (int page = 1; page <= 2; page++) {
            bool tablePage = variant == "default" || variant == "first" && page == 1 || variant == "even" && page == 2;
            Assert.Equal(tablePage && image ? 1 : 0, pdf.GetPage(page).GetImages().Count());
            if (variant == "default" || !tablePage)
                Assert.Contains($"NUMBER {page}/2", pdf.GetPage(page).Text);
            else Assert.DoesNotContain("NUMBER", pdf.GetPage(page).Text);
        }
        Assert.Equal(before, string.Concat(story.ChildElements.Select(element => element.OuterXml)));
    }
}
