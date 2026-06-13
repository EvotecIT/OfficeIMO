using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word.Pdf;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.Net.Http;
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
            MaxTotalCssBytes = 16384,
            MaxTableCells = 10000,
            EnableAccessibilityDiagnostics = true
        };
        Assert.Equal("Calibri", options.FontFamily);
        Assert.Equal(1024, options.MaxImageBytes);
        Assert.Equal(4096, options.MaxTotalImageBytes);
        Assert.Equal(1000, options.MaxHtmlNodes);
        Assert.Equal(64, options.MaxHtmlDepth);
        Assert.Equal(8192, options.MaxCssBytes);
        Assert.Equal(16384, options.MaxTotalCssBytes);
        Assert.Equal(10000, options.MaxTableCells);
        Assert.False(options.AllowDocumentStylesheetLinks);
        Assert.True(options.ValidateImageContentTypes);
        Assert.Contains("image/png", options.AllowedImageContentTypes);
        Assert.Contains("https", options.AllowedImageUriSchemes);
        Assert.Empty(options.AllowedImageHosts);
        Assert.True(options.ValidateStylesheetContentTypes);
        Assert.Contains("text/css", options.AllowedStylesheetContentTypes);
        Assert.Contains("https", options.AllowedStylesheetUriSchemes);
        Assert.Contains("file", options.AllowedStylesheetUriSchemes);
        Assert.Empty(options.AllowedStylesheetHosts);
        Assert.True(options.EnableAccessibilityDiagnostics);
    }

    [Fact]
    public void HtmlToWordOptions_CreateProfilesExposeExpectedPolicies() {
        var defaultProfile = HtmlToWordOptions.CreateOfficeIMOProfile();
        Assert.False(defaultProfile.AllowDocumentStylesheetLinks);
        Assert.Equal(ImageProcessingMode.Embed, defaultProfile.ImageProcessing);
        Assert.Contains("https", defaultProfile.AllowedImageUriSchemes);
        Assert.Contains("https", defaultProfile.AllowedStylesheetUriSchemes);
        Assert.Null(defaultProfile.MaxHtmlNodes);

        var untrustedProfile = HtmlToWordOptions.CreateUntrustedHtmlProfile();
        Assert.False(untrustedProfile.AllowDocumentStylesheetLinks);
        Assert.Equal(ImageProcessingMode.EmbedDataUriOnly, untrustedProfile.ImageProcessing);
        Assert.Equal(TimeSpan.FromSeconds(5), untrustedProfile.ResourceTimeout);
        Assert.Equal(10000, untrustedProfile.MaxHtmlNodes);
        Assert.Equal(64, untrustedProfile.MaxHtmlDepth);
        Assert.Equal(256L * 1024L, untrustedProfile.MaxCssBytes);
        Assert.Equal(512L * 1024L, untrustedProfile.MaxTotalCssBytes);
        Assert.Equal(50000, untrustedProfile.MaxTableCells);
        Assert.Equal(new[] { "data" }, untrustedProfile.AllowedImageUriSchemes);
        Assert.Empty(untrustedProfile.AllowedStylesheetUriSchemes);
        Assert.True(untrustedProfile.EnableAccessibilityDiagnostics);

        var trustedProfile = HtmlToWordOptions.CreateTrustedDocumentProfile();
        Assert.True(trustedProfile.AllowDocumentStylesheetLinks);
        Assert.Equal(ImageProcessingMode.Embed, trustedProfile.ImageProcessing);
        Assert.True(trustedProfile.ValidateImageContentTypes);
        Assert.True(trustedProfile.ValidateStylesheetContentTypes);
    }

    [Fact]
    public void HtmlToWordOptions_CloneCopiesConfigurationWithoutDiagnostics() {
        using var httpClient = new HttpClient();
        Action<HtmlConversionDiagnostic> handler = _ => { };
        var options = HtmlToWordOptions.CreateTrustedDocumentProfile();
        options.FontFamily = "Aptos";
        options.QuotePrefix = "[";
        options.QuoteSuffix = "]";
        options.DefaultPageSize = WordPageSize.A4;
        options.DefaultOrientation = PageOrientationValues.Landscape;
        options.IncludeListStyles = true;
        options.ContinueNumbering = true;
        options.SupportsHeadingNumbering = true;
        options.BasePath = "C:\\Temp";
        options.NoteReferenceType = NoteReferenceType.Endnote;
        options.LinkNoteUrls = false;
        options.ImageProcessing = ImageProcessingMode.LinkExternal;
        options.HttpClient = httpClient;
        options.ResourceTimeout = TimeSpan.FromSeconds(7);
        options.MaxImageBytes = 11;
        options.MaxTotalImageBytes = 22;
        options.ValidateImageContentTypes = false;
        options.ValidateStylesheetContentTypes = false;
        options.MaxHtmlNodes = 33;
        options.MaxHtmlDepth = 44;
        options.MaxCssBytes = 55;
        options.MaxTotalCssBytes = 66;
        options.MaxTableCells = 77;
        options.DiagnosticHandler = handler;
        options.EnableAccessibilityDiagnostics = true;
        options.ImportHtmlComments = true;
        options.HtmlCommentAuthor = "HTML Reviewer";
        options.HtmlCommentInitials = "HR";
        options.UnsupportedCssHandling = HtmlUnsupportedCssHandling.Error;
        options.RenderPreAsTable = true;
        options.TableCaptionPosition = TableCaptionPosition.Below;
        options.SectionTagHandling = SectionTagHandling.Block;
        options.ClassStyles["lead"] = WordParagraphStyles.Heading2;
        options.StylesheetPaths.Add("site.css");
        options.StylesheetContents.Add("p{color:red}");
        options.AllowedImageContentTypes.Clear();
        options.AllowedImageContentTypes.Add("image/png");
        options.AllowedImageUriSchemes.Clear();
        options.AllowedImageUriSchemes.Add("data");
        options.AllowedImageHosts.Add("images.example.test");
        options.AllowedStylesheetUriSchemes.Clear();
        options.AllowedStylesheetUriSchemes.Add("https");
        options.AllowedStylesheetHosts.Add("styles.example.test");
        options.AllowedStylesheetContentTypes.Clear();
        options.AllowedStylesheetContentTypes.Add("text/css");
        options.Diagnostics.Add(new HtmlConversionDiagnostic("Existing", "Existing diagnostic"));

        var clone = options.Clone();

        Assert.Equal(options.FontFamily, clone.FontFamily);
        Assert.Equal(options.QuotePrefix, clone.QuotePrefix);
        Assert.Equal(options.QuoteSuffix, clone.QuoteSuffix);
        Assert.Equal(options.DefaultPageSize, clone.DefaultPageSize);
        Assert.Equal(options.DefaultOrientation, clone.DefaultOrientation);
        Assert.Equal(options.IncludeListStyles, clone.IncludeListStyles);
        Assert.Equal(options.ContinueNumbering, clone.ContinueNumbering);
        Assert.Equal(options.SupportsHeadingNumbering, clone.SupportsHeadingNumbering);
        Assert.Equal(options.BasePath, clone.BasePath);
        Assert.Equal(options.NoteReferenceType, clone.NoteReferenceType);
        Assert.Equal(options.LinkNoteUrls, clone.LinkNoteUrls);
        Assert.Equal(options.ImageProcessing, clone.ImageProcessing);
        Assert.Same(httpClient, clone.HttpClient);
        Assert.Equal(options.ResourceTimeout, clone.ResourceTimeout);
        Assert.Equal(options.MaxImageBytes, clone.MaxImageBytes);
        Assert.Equal(options.MaxTotalImageBytes, clone.MaxTotalImageBytes);
        Assert.Equal(options.ValidateImageContentTypes, clone.ValidateImageContentTypes);
        Assert.Equal(options.ValidateStylesheetContentTypes, clone.ValidateStylesheetContentTypes);
        Assert.Equal(options.MaxHtmlNodes, clone.MaxHtmlNodes);
        Assert.Equal(options.MaxHtmlDepth, clone.MaxHtmlDepth);
        Assert.Equal(options.MaxCssBytes, clone.MaxCssBytes);
        Assert.Equal(options.MaxTotalCssBytes, clone.MaxTotalCssBytes);
        Assert.Equal(options.MaxTableCells, clone.MaxTableCells);
        Assert.Same(handler, clone.DiagnosticHandler);
        Assert.Equal(options.EnableAccessibilityDiagnostics, clone.EnableAccessibilityDiagnostics);
        Assert.Equal(options.ImportHtmlComments, clone.ImportHtmlComments);
        Assert.Equal(options.HtmlCommentAuthor, clone.HtmlCommentAuthor);
        Assert.Equal(options.HtmlCommentInitials, clone.HtmlCommentInitials);
        Assert.Equal(options.UnsupportedCssHandling, clone.UnsupportedCssHandling);
        Assert.Equal(options.AllowDocumentStylesheetLinks, clone.AllowDocumentStylesheetLinks);
        Assert.Equal(options.RenderPreAsTable, clone.RenderPreAsTable);
        Assert.Equal(options.TableCaptionPosition, clone.TableCaptionPosition);
        Assert.Equal(options.SectionTagHandling, clone.SectionTagHandling);
        Assert.Equal(options.ClassStyles, clone.ClassStyles);
        Assert.Equal(options.StylesheetPaths, clone.StylesheetPaths);
        Assert.Equal(options.StylesheetContents, clone.StylesheetContents);
        Assert.Equal(options.AllowedImageContentTypes, clone.AllowedImageContentTypes);
        Assert.Equal(options.AllowedImageUriSchemes, clone.AllowedImageUriSchemes);
        Assert.Equal(options.AllowedImageHosts, clone.AllowedImageHosts);
        Assert.Equal(options.AllowedStylesheetUriSchemes, clone.AllowedStylesheetUriSchemes);
        Assert.Equal(options.AllowedStylesheetHosts, clone.AllowedStylesheetHosts);
        Assert.Equal(options.AllowedStylesheetContentTypes, clone.AllowedStylesheetContentTypes);
        Assert.Empty(clone.Diagnostics);

        clone.ClassStyles["lead"] = WordParagraphStyles.Heading3;
        clone.StylesheetPaths.Add("other.css");
        clone.AllowedImageUriSchemes.Add("https");

        Assert.Equal(WordParagraphStyles.Heading2, options.ClassStyles["lead"]);
        Assert.DoesNotContain("other.css", options.StylesheetPaths);
        Assert.DoesNotContain("https", options.AllowedImageUriSchemes);
    }

    [Fact]
    public void WordToMarkdownOptions_ExposeFontFamily() {
        var options = new WordToMarkdownOptions {
            FontFamily = "Arial",
            PageBreakMode = MarkdownPageBreakMode.HorizontalRule,
            UnsupportedContentMode = MarkdownUnsupportedContentMode.Placeholder,
            VisualFallbackMode = MarkdownVisualFallbackMode.SvgFile,
            VisualFallbackDirectory = "assets",
            VisualFallbackPathPrefix = "assets"
        };
        Assert.Equal("Arial", options.FontFamily);
        Assert.Equal(MarkdownPageBreakMode.HorizontalRule, options.PageBreakMode);
        Assert.Equal(MarkdownUnsupportedContentMode.Placeholder, options.UnsupportedContentMode);
        Assert.Equal(MarkdownVisualFallbackMode.SvgFile, options.VisualFallbackMode);
        Assert.Equal("assets", options.VisualFallbackDirectory);
        Assert.Equal("assets", options.VisualFallbackPathPrefix);
    }

    [Fact]
    public void MarkdownToWordOptions_EnableBoundedDataUriImagesByDefault() {
        var options = new MarkdownToWordOptions();

        Assert.True(options.AllowDataUriImages);
        Assert.Equal(32L * 1024L * 1024L, options.MaxDataUriImageBytes);
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
