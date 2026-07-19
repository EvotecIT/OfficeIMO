using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Epub;
using OfficeIMO.Epub.Image;
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
    public void RetainedEpubResourcesRenderSynchronously() {
        using var package = new MemoryStream(CreateEpub());
        EpubDocument book = EpubDocument.Load(
            package,
            new EpubReadOptions {
                IncludeRawHtml = true,
                IncludeResourceData = true
            });

        OfficeImageExportResult result = Assert.Single(
            book.ExportImages(OfficeImageExportFormat.Png));

        Assert.DoesNotContain(
            result.Diagnostics,
            diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ExternalImagePending ||
                          diagnostic.Code == HtmlRenderDiagnosticCodes.ExternalStylesheetPending ||
                          diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceUnavailable ||
                          diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task EpubRetainedResourceByteLimitHasPreciseDiagnostic(
        bool asynchronous) {
        using var package = new MemoryStream(CreateEpub());
        EpubDocument book = EpubDocument.Load(
            package,
            new EpubReadOptions {
                IncludeRawHtml = true,
                IncludeResourceData = true
            });
        var options = new EpubImageExportOptions {
            MaxResourceBytes = PixelPng.Length - 1L
        };

        IReadOnlyList<OfficeImageExportResult> results = asynchronous
            ? await book.ExportImagesAsync(
                OfficeImageExportFormat.Png,
                options)
            : book.ExportImages(
                OfficeImageExportFormat.Png,
                options);

        OfficeImageExportResult result = Assert.Single(results);
        Assert.Contains(
            result.Diagnostics,
            diagnostic => diagnostic.Code ==
                          HtmlRenderDiagnosticCodes.ResourceByteLimitExceeded);
        Assert.DoesNotContain(
            result.Diagnostics,
            diagnostic => diagnostic.Code ==
                          HtmlRenderDiagnosticCodes.ResourceUnavailable);
    }

    [Fact]
    public void EpubPackageResourcePolicyDoesNotAuthorizeHyperlinks() {
        using var package = new MemoryStream(CreateEpub());
        EpubDocument book = EpubDocument.Load(
            package,
            new EpubReadOptions {
                IncludeRawHtml = true,
                IncludeResourceData = true
            });
        var options = new EpubImageExportOptions {
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
        };

        OfficeImageExportResult result = Assert.Single(
            book.ExportImages(
                OfficeImageExportFormat.Svg,
                options));

        Assert.Contains(
            result.Diagnostics,
            diagnostic => diagnostic.Code == "HyperlinkRejectedByPolicy");
        Assert.DoesNotContain(
            "epub://",
            Encoding.UTF8.GetString(result.Bytes),
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task MissingEpubPackageResourceDoesNotReachExternalFallback() {
        using var package = new MemoryStream(
            CreateEpub(includeImage: false));
        EpubDocument book = EpubDocument.Load(
            package,
            new EpubReadOptions {
                IncludeRawHtml = true,
                IncludeResourceData = true
            });
        int fallbackCalls = 0;
        var options = new EpubImageExportOptions {
            ResourceResolver = (request, cancellationToken) => {
                cancellationToken.ThrowIfCancellationRequested();
                fallbackCalls++;
                return Task.FromResult<HtmlResolvedResource?>(
                    new HtmlResolvedResource(PixelPng, "image/png"));
            }
        };

        OfficeImageExportResult result = Assert.Single(
            await book.ExportImagesAsync(
                OfficeImageExportFormat.Png,
                options));

        Assert.Equal(0, fallbackCalls);
        Assert.Contains(
            result.Diagnostics,
            diagnostic => diagnostic.Code ==
                          HtmlRenderDiagnosticCodes.ResourceUnavailable);
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

    [Fact]
    public void EpubRetainedPackageWarningsDoNotMasqueradeAsOmissions() {
        using var package = new MemoryStream(
            CreateEpub(includePackageVersion: false));
        EpubDocument book = EpubDocument.Load(
            package,
            new EpubReadOptions {
                IncludeRawHtml = true,
                IncludeResourceData = true
            });
        var options = new EpubImageExportOptions {
            Policy = new OfficeImageExportPolicy { RequireNoOmissions = true }
        };

        OfficeImageExportResult result = Assert.Single(
            book.ExportImages(OfficeImageExportFormat.Png, options));

        Assert.Contains(
            result.Diagnostics,
            diagnostic =>
                diagnostic.Code == "EPUB_IMAGE_EPUB_PACKAGE_VERSION_MISSING" &&
                diagnostic.LossKind == OfficeImageExportLossKind.Approximation);
    }

    [Fact]
    public void EpubRawHtmlLimitFallbackParticipatesInNoOmissionsPolicy() {
        using var package = new MemoryStream(CreateEpub());
        EpubDocument book = EpubDocument.Load(
            package,
            new EpubReadOptions {
                IncludeRawHtml = true,
                IncludeResourceData = true,
                MaxTotalRawHtmlBytes = 1
            });
        var options = new EpubImageExportOptions {
            Policy = new OfficeImageExportPolicy { RequireNoOmissions = true }
        };

        OfficeImageExportPolicyException exception = Assert.Throws<OfficeImageExportPolicyException>(() =>
            book.ExportImages(OfficeImageExportFormat.Png, options));

        Assert.Contains(
            exception.Diagnostics,
            diagnostic =>
                diagnostic.Code == "EPUB_IMAGE_EPUB_CHAPTER_RAW_HTML_TOTAL_LIMIT" &&
                diagnostic.LossKind == OfficeImageExportLossKind.Omission);
    }

    private static byte[] CreateEpub(
        bool includeImage = true,
        bool includePackageVersion = true) {
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
                "<?xml version=\"1.0\"?><package xmlns=\"http://www.idpf.org/2007/opf\"" +
                (includePackageVersion ? " version=\"3.0\"" : string.Empty) +
                " unique-identifier=\"id\"><metadata xmlns:dc=\"http://purl.org/dc/elements/1.1/\"><dc:identifier id=\"id\">book</dc:identifier><dc:title>Image Book</dc:title></metadata><manifest><item id=\"chapter\" href=\"chapter.xhtml\" media-type=\"application/xhtml+xml\"/><item id=\"styles\" href=\"styles/book.css\" media-type=\"text/css\"/><item id=\"pixel\" href=\"images/pixel.png\" media-type=\"image/png\"/></manifest><spine><itemref idref=\"chapter\"/></spine></package>");
            Write(
                archive,
                "OEBPS/chapter.xhtml",
                "<?xml version=\"1.0\"?><html xmlns=\"http://www.w3.org/1999/xhtml\"><head><title>Chapter One</title><link rel=\"stylesheet\" href=\"styles/book.css\" /></head><body><h1>Chapter One</h1><p>Rendered EPUB content</p><a href=\"next.xhtml\">Next chapter</a><img src=\"images/pixel.png\" alt=\"pixel\"/></body></html>");
            Write(
                archive,
                "OEBPS/styles/book.css",
                "body { color: #123456; }");
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
