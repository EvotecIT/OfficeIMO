using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Epub;
using OfficeIMO.Epub.Html;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlEpubImageExportTests {
    private static readonly byte[] PixelPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAF/gL+X8m0WQAAAABJRU5ErkJggg==");

    [Fact]
    public async Task RetainedEpubChapterAndResourcesRenderThroughHtml() {
        using var package = new MemoryStream(CreateEpub());
        EpubDocument book = EpubDocument.Load(
            package,
            new EpubReadOptions {
                IncludeRawHtml = true,
                IncludeResourceData = true
            });

        var results = await book
            .ToImages()
            .Continuous()
            .AsPng()
            .ExportAsync();

        OfficeImageExportResult result = Assert.Single(results);
        Assert.Equal("Chapter One", result.Name);
        Assert.Equal("OEBPS/chapter.xhtml", result.Source);
        Assert.DoesNotContain(
            result.Diagnostics,
            diagnostic => diagnostic.Code ==
                          HtmlRenderDiagnosticCodes.ResourceUnavailable);
        Assert.DoesNotContain(
            result.Diagnostics,
            diagnostic => diagnostic.Code ==
                          OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
    }

    [Fact]
    public void EpubWithoutRawHtmlUsesDiagnosedTextFallback() {
        using var package = new MemoryStream(CreateEpub());
        EpubDocument book = EpubDocument.Load(package);

        OfficeImageExportResult result = Assert.Single(
            book.ExportImages(OfficeImageExportFormat.Svg));

        Assert.Contains(
            result.Diagnostics,
            diagnostic => diagnostic.Code ==
                          "EPUB_IMAGE_RAW_HTML_UNAVAILABLE" &&
                          diagnostic.LossKind ==
                          OfficeImageExportLossKind.Approximation);
    }

    [Fact]
    public void EpubTextFallbackParticipatesInNoLossPolicy() {
        using var package = new MemoryStream(CreateEpub());
        EpubDocument book = EpubDocument.Load(package);
        var options = new EpubImageExportOptions {
            Policy = new OfficeImageExportPolicy { RequireNoLoss = true }
        };

        Assert.Throws<OfficeImageExportPolicyException>(() =>
            book.ExportImages(
                OfficeImageExportFormat.Png,
                options));
    }

    [Fact]
    public void EpubPackageOmissionsParticipateInNoOmissionsPolicy() {
        using var package = new MemoryStream(CreateEpub(includeImage: false));
        EpubDocument book = EpubDocument.Load(
            package,
            new EpubReadOptions {
                IncludeRawHtml = true,
                IncludeResourceData = true
            });
        var options = new EpubImageExportOptions {
            Policy = new OfficeImageExportPolicy { RequireNoOmissions = true }
        };

        OfficeImageExportPolicyException exception = Assert.Throws<OfficeImageExportPolicyException>(() =>
            book.ExportImages(OfficeImageExportFormat.Png, options));

        Assert.Contains(
            exception.Diagnostics,
            diagnostic =>
                diagnostic.Code == "EPUB_IMAGE_EPUB_RESOURCE_MISSING" &&
                diagnostic.LossKind == OfficeImageExportLossKind.Omission);
    }

    private static byte[] CreateEpub(bool includeImage = true) {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(
                   output,
                   ZipArchiveMode.Create,
                   leaveOpen: true)) {
            Write(
                archive,
                "mimetype",
                "application/epub+zip",
                CompressionLevel.NoCompression);
            Write(
                archive,
                "META-INF/container.xml",
                "<?xml version=\"1.0\"?><container xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\" version=\"1.0\"><rootfiles><rootfile full-path=\"OEBPS/content.opf\" media-type=\"application/oebps-package+xml\"/></rootfiles></container>");
            Write(
                archive,
                "OEBPS/content.opf",
                "<?xml version=\"1.0\"?><package xmlns=\"http://www.idpf.org/2007/opf\" version=\"3.0\" unique-identifier=\"id\"><metadata xmlns:dc=\"http://purl.org/dc/elements/1.1/\"><dc:identifier id=\"id\">book</dc:identifier><dc:title>Image Book</dc:title></metadata><manifest><item id=\"chapter\" href=\"chapter.xhtml\" media-type=\"application/xhtml+xml\"/><item id=\"pixel\" href=\"images/pixel.png\" media-type=\"image/png\"/></manifest><spine><itemref idref=\"chapter\"/></spine></package>");
            Write(
                archive,
                "OEBPS/chapter.xhtml",
                "<?xml version=\"1.0\"?><html xmlns=\"http://www.w3.org/1999/xhtml\"><head><title>Chapter One</title></head><body><h1>Chapter One</h1><p>Rendered EPUB content</p><img src=\"images/pixel.png\" alt=\"pixel\"/></body></html>");
            if (includeImage) {
                ZipArchiveEntry image = archive.CreateEntry(
                    "OEBPS/images/pixel.png");
                using Stream stream = image.Open();
                stream.Write(PixelPng, 0, PixelPng.Length);
            }
        }
        return output.ToArray();
    }

    private static void Write(
        ZipArchive archive,
        string path,
        string value,
        CompressionLevel compression = CompressionLevel.Optimal) {
        ZipArchiveEntry entry = archive.CreateEntry(path, compression);
        using Stream stream = entry.Open();
        byte[] bytes = Encoding.UTF8.GetBytes(value);
        stream.Write(bytes, 0, bytes.Length);
    }
}
