using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Pdf;
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
    public void PdfSaveOptions_ExposeFontFamily() {
        var options = new PdfSaveOptions { FontFamily = "Times New Roman" };
        Assert.Equal("Times New Roman", options.FontFamily);
    }
}
