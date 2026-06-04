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
        var options = new HtmlToWordOptions {
            FontFamily = "Calibri",
            MaxImageBytes = 1024,
            MaxTotalImageBytes = 4096,
            MaxHtmlNodes = 1000,
            MaxHtmlDepth = 64,
            MaxCssBytes = 8192,
            MaxTableCells = 10000,
            EnableAccessibilityDiagnostics = true
        };
        Assert.Equal("Calibri", options.FontFamily);
        Assert.Equal(1024, options.MaxImageBytes);
        Assert.Equal(4096, options.MaxTotalImageBytes);
        Assert.Equal(1000, options.MaxHtmlNodes);
        Assert.Equal(64, options.MaxHtmlDepth);
        Assert.Equal(8192, options.MaxCssBytes);
        Assert.Equal(10000, options.MaxTableCells);
        Assert.True(options.ValidateImageContentTypes);
        Assert.Contains("image/png", options.AllowedImageContentTypes);
        Assert.Contains("https", options.AllowedImageUriSchemes);
        Assert.Empty(options.AllowedImageHosts);
        Assert.True(options.EnableAccessibilityDiagnostics);
    }

    [Fact]
    public void WordToMarkdownOptions_ExposeFontFamily() {
        var options = new WordToMarkdownOptions { FontFamily = "Arial" };
        Assert.Equal("Arial", options.FontFamily);
    }

    [Fact]
    public void WordToHtmlOptions_ExposeSectionMetadataOption() {
        var options = new WordToHtmlOptions {
            IncludeCustomProperties = true,
            IncludeSectionMetadata = true
        };
        Assert.True(options.IncludeCustomProperties);
        Assert.True(options.IncludeSectionMetadata);
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
