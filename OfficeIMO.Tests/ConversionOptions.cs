using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word.Pdf;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public class ConversionOptionsTests {
    [Fact]
    public void HtmlToWordOptions_ExposeFontFamily() {
        var options = new HtmlToWordOptions { FontFamily = "Calibri" };
        Assert.Equal("Calibri", options.FontFamily);
    }

    [Fact]
    public void WordToMarkdownOptions_ExposeFontFamily() {
        var options = new WordToMarkdownOptions { FontFamily = "Arial" };
        Assert.Equal("Arial", options.FontFamily);
    }

    [Fact]
    public void PdfOptions_ExposeFontFamily() {
        var options = new PdfSaveOptions { FontFamily = "Times New Roman" };
        Assert.Equal("Times New Roman", options.FontFamily);
    }

    [Fact]
    public void Options_ExposeDefaultPageSettings() {
        var options = new HtmlToWordOptions {
            DefaultOrientation = PageOrientationValues.Landscape,
            DefaultPageSize = WordPageSize.A3
        };
        Assert.Equal(PageOrientationValues.Landscape, options.DefaultOrientation);
        Assert.Equal(WordPageSize.A3, options.DefaultPageSize);
    }
}
