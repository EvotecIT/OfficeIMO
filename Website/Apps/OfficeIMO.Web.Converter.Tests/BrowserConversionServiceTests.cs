using System.IO.Compression;
using OfficeIMO.Drawing;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Web.Converter.Services;
using Xunit;

namespace OfficeIMO.Web.Converter.Tests;

public sealed class BrowserConversionServiceTests {
    private readonly BrowserConversionService _service = new();

    [Fact]
    public void RouteCatalog_HasUniqueCustomerRoutes() {
        Assert.Equal(6, ConversionRouteCatalog.All.Count);
        Assert.Equal(
            ConversionRouteCatalog.All.Count,
            ConversionRouteCatalog.All.Select(static route => route.Id).Distinct(StringComparer.OrdinalIgnoreCase).Count());
        Assert.All(ConversionRouteCatalog.All, static route => {
            Assert.False(string.IsNullOrWhiteSpace(route.Source));
            Assert.False(string.IsNullOrWhiteSpace(route.Target));
            Assert.False(string.IsNullOrWhiteSpace(route.EnginePath));
        });
    }

    [Fact]
    public void MarkdownToHtml_ReturnsPreviewAndDownload() {
        var route = ConversionRouteCatalog.Find("markdown-html");
        var result = _service.ConvertText(route, "# Status\n\n**Ready**");

        Assert.Equal("text/html;charset=utf-8", result.ContentType);
        Assert.Equal("officeimo-markdown.html", result.FileName);
        Assert.Contains("<h1", result.HtmlPreview, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Ready", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_UsesSharedHtmlDocument() {
        var route = ConversionRouteCatalog.Find("html-markdown");
        var result = _service.ConvertText(route, "<h1>Status</h1><p>Ready</p>");

        Assert.Equal("text/markdown;charset=utf-8", result.ContentType);
        Assert.Contains("# Status", result.Text, StringComparison.Ordinal);
        Assert.Contains("Ready", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownToWord_ReturnsOpenXmlPackage() {
        var route = ConversionRouteCatalog.Find("markdown-docx");
        var result = _service.ConvertText(route, "# Status\n\nReady");

        Assert.Equal("officeimo-markdown.docx", result.FileName);
        Assert.True(result.Bytes.Length > 4);
        Assert.Equal((byte)'P', result.Bytes[0]);
        Assert.Equal((byte)'K', result.Bytes[1]);
    }

    [Fact]
    public void TextConversion_RejectsInputBeyondBrowserLimit() {
        var route = ConversionRouteCatalog.Find("markdown-html");
        string oversized = new('a', BrowserConversionService.MaxTextInputChars + 1);

        var exception = Assert.Throws<ArgumentOutOfRangeException>(() => _service.ConvertText(route, oversized));

        Assert.Contains("browser converter", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void WordConversion_RejectsCompressedPackageBomb() {
        var route = ConversionRouteCatalog.Find("docx-pdf");
        byte[] bytes = CreateHighlyCompressedPackage();
        var document = new SelectedDocument("unsafe.docx", ".docx", "DOCX", bytes.LongLength, bytes);

        Assert.Throws<OfficePackageSecurityException>(() => _service.ConvertFile(route, document, fastPreview: false));
    }

    private static byte[] CreateHighlyCompressedPackage() {
        using var buffer = new MemoryStream();
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true)) {
            ZipArchiveEntry contentTypes = archive.CreateEntry("[Content_Types].xml", CompressionLevel.Optimal);
            using (var writer = new StreamWriter(contentTypes.Open())) {
                writer.Write("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\" />");
            }

            ZipArchiveEntry oversizedPart = archive.CreateEntry("word/document.xml", CompressionLevel.Optimal);
            using var stream = oversizedPart.Open();
            byte[] repeated = new byte[2 * 1024 * 1024];
            Array.Fill(repeated, (byte)'A');
            stream.Write(repeated);
        }
        return buffer.ToArray();
    }
}
