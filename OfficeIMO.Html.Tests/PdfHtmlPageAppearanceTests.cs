using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using AngleSharp.Html.Parser;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Tests.Pdf;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PdfHtmlPageAppearanceTests {
    [Fact]
    public void DetailedImageFallsBackWithoutFailingConversion() {
        var raster = new OfficeRasterImage(1200, 1200, OfficeColor.White);
        for (int y = 0; y < raster.Height; y++)
            for (int x = y % 2; x < raster.Width; x += 2) raster.SetPixel(x, y, OfficeColor.Black);
        byte[] image = OfficePngWriter.Encode(raster);
        var pdf = PdfDocument.Load(PdfDocument.Create()
            .Canvas(canvas => canvas.Image(image, 40, 50, 300, 300)).ToBytes());
        var result = pdf.ToHtmlResult(PdfToHtmlOptions.CreatePositionedReviewProfile());
        Assert.DoesNotContain("pdf-page-appearance", result.Value);
        Assert.Contains("data:image/png;base64,", result.Value);
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "PageAppearanceImageFallback");
    }

    [Fact]
    public void OpenedPdfHonorsImageEmbeddingChoicesAndLimits() {
        var pdf = PdfDocument.Load(PdfDocument.Create()
            .Canvas(canvas => canvas.Image(PdfPngTestImages.CreateRgbPng(2, 2), 40, 50, 60, 30)).ToBytes());
        var options = PdfToHtmlOptions.CreatePositionedReviewProfile();
        string embedded = pdf.ToHtml(options);
        Assert.Contains("pdf-page-appearance", embedded);
        using var imageHtml = new HtmlParser().ParseDocument(embedded);
        Assert.NotEmpty(imageHtml.QuerySelectorAll(".pdf-page-appearance svg rect[fill]"));
        Assert.Equal(0, pdf.ToHtmlResult(options).Summary.ImagePlaceholderCount);
        options.MaxEmbeddedImageBytes = 0;
        var limited = pdf.ToHtmlResult(options);
        Assert.DoesNotContain("pdf-page-appearance", limited.Value);
        Assert.DoesNotContain("data:image/png;base64,", limited.Value);
        Assert.Equal(1, limited.Summary.ImagePlaceholderCount);
        Assert.Contains(limited.Report.Warnings, warning => warning.Code == "ImageDataTooLarge");
        options.MaxEmbeddedImageBytes = null;
        options.ImageExportMode = PdfHtmlImageExportMode.PlaceholderOnly;
        Assert.DoesNotContain("data:image/png;base64,", pdf.ToHtml(options));
        options.ImageExportMode = PdfHtmlImageExportMode.EmbeddedDataUri;
        options.IncludeImagePlaceholders = false;
        Assert.DoesNotContain("data:image/png;base64,", pdf.ToHtml(options));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnsupportedPageAppearanceImageFallsBackToLogicalImage(bool interpolate) {
        byte[] jpeg2000 = File.ReadAllBytes(Path.Combine(
            AppContext.BaseDirectory, "Pdf", "Fixtures", "Interoperability", "Scans", "red-rgb.jp2"));
        var pdf = PdfDocument.Load(BuildImagePdf(jpeg2000, "/JPXDecode", interpolate));

        var result = pdf.ToHtmlResult(PdfToHtmlOptions.CreatePositionedReviewProfile());

        Assert.DoesNotContain("pdf-page-appearance", result.Value);
        Assert.Contains("data:image/jp2;base64,", result.Value);
        Assert.Equal(1, result.Summary.ImagePlaceholderCount);
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "PageAppearanceImageFallback");
    }

    [Fact]
    public void UndecodableNearestNeighborJpegFallsBackWithoutFailingConversion() {
        var pdf = PdfDocument.Load(BuildImagePdf(new byte[] { 0xff, 0xd8, 0xff, 0xd9 }, "/DCTDecode", interpolate: false));

        var result = pdf.ToHtmlResult(PdfToHtmlOptions.CreatePositionedReviewProfile());

        Assert.DoesNotContain("pdf-page-appearance", result.Value);
        Assert.Contains("data:image/jpeg;base64,", result.Value);
        Assert.Equal(1, result.Summary.ImagePlaceholderCount);
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "PageAppearanceImageFallback");
    }

    [Fact]
    public async Task OpenedPdfKeepsArtworkAndSelectableTextAcrossOutputEntrypoints() {
        var pdf = PdfDocument.Load(Source());
        var options = PdfToHtmlOptions.CreatePositionedReviewProfile();
        var result = pdf.ToHtmlResult(options);
        using var html = new HtmlParser().ParseDocument(result.Value);
        Assert.Single(html.QuerySelectorAll(".pdf-page-appearance svg"));
        Assert.NotEmpty(html.QuerySelectorAll(".pdf-page-appearance svg path, .pdf-page-appearance svg rect"));
        Assert.Contains("Invoice sample", string.Join(" ", html.QuerySelectorAll("svg text").Select(node => node.TextContent)));
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "VectorAppearanceNotExported");
        Assert.Equal(result.Value, pdf.ToHtml(options));
        using var stream = new MemoryStream();
        pdf.SaveAsHtml(stream, options);
        Assert.Equal(result.Value, Encoding.UTF8.GetString(stream.ToArray()));
        using var asyncStream = new MemoryStream();
        await pdf.SaveAsHtmlAsync(asyncStream, options);
        Assert.Equal(result.Value, Encoding.UTF8.GetString(asyncStream.ToArray()));
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_HTML_VISUAL_OUTPUT");
        if (!string.IsNullOrWhiteSpace(output)) {
            Directory.CreateDirectory(output);
            File.WriteAllText(Path.Combine(output, "page-appearance.html"), result.Value);
            File.WriteAllBytes(Path.Combine(output, "page-appearance.pdf"), Source());
        }
    }

    [Fact]
    public void AppearanceKeepsSourcePageSelectionAndRotatedGeometry() {
        var pdf = PdfDocument.Load(Source(twoPages: true));
        var options = PdfToHtmlOptions.CreatePositionedReviewProfile();
        options.PageRanges = new[] { new PdfPageRange(2, 2) };
        var result = pdf.ToHtmlResult(options);
        using var html = new HtmlParser().ParseDocument(result.Value);
        var page = Assert.Single(html.QuerySelectorAll("section.pdf-page"));
        Assert.Equal("2", page.GetAttribute("data-page-number"));
        Assert.Equal("width:180pt;height:240pt;", page.GetAttribute("style"));
        Assert.Contains("Second page", string.Join(" ", page.QuerySelectorAll("svg text").Select(node => node.TextContent)));
        Assert.DoesNotContain("Invoice sample", string.Join(" ", page.QuerySelectorAll("svg text").Select(node => node.TextContent)));
        Assert.Equal(new[] { 2 }, result.Summary.PageNumbers);
        Assert.Equal(2, result.Summary.SourcePageCount);
        Assert.Single(options.PageRanges);
    }

    [Fact]
    public void AppearanceHonorsCancellationAndOutputBudget() {
        var pdf = PdfDocument.Load(Source());
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => pdf.ToHtmlResult(PdfToHtmlOptions.CreatePositionedReviewProfile(), cancellation.Token));
        var options = PdfToHtmlOptions.CreatePositionedReviewProfile();
        string complete = pdf.ToHtml(options);
        options.MaximumOutputCharacters = complete.Length;
        Assert.Equal(complete, pdf.ToHtml(options));
        options.MaximumOutputCharacters = complete.Length - 100;
        Assert.Throws<InvalidOperationException>(() => pdf.ToHtmlResult(options));
    }

    private static byte[] Source(bool twoPages = false) {
        const string content = "0.1 0.5 0.9 rg 40 70 140 45 re f 0 0 0 rg BT /F1 16 Tf 40 140 Td (Invoice sample) Tj ET";
        string second = content.Replace("Invoice sample", "Second page");
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7", "1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj",
            twoPages ? "2 0 obj << /Type /Pages /Count 2 /Kids [3 0 R 6 0 R] >> endobj" : "2 0 obj << /Type /Pages /Count 1 /Kids [3 0 R] >> endobj",
            "3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >> endobj",
            $"4 0 obj << /Length {content.Length} >> stream", content, "endstream endobj",
            "5 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj",
            twoPages ? "6 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Rotate 90 /Resources << /Font << /F1 5 0 R >> >> /Contents 7 0 R >> endobj" : "",
            twoPages ? $"7 0 obj << /Length {second.Length} >> stream\n" + second + "\nendstream endobj" : "",
            "trailer << /Root 1 0 R /Size 8 >>", "%%EOF"
        }));
    }

    private static byte[] BuildImagePdf(byte[] image, string filter, bool interpolate) {
        byte[] content = Encoding.ASCII.GetBytes("q 80 0 0 60 40 50 cm /Im1 Do Q");
        using var pdf = new MemoryStream();
        WriteAscii(pdf, "%PDF-1.7\n");
        WriteAscii(pdf, "1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj\n");
        WriteAscii(pdf, "2 0 obj << /Type /Pages /Count 1 /Kids [3 0 R] >> endobj\n");
        WriteAscii(pdf, "3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >> endobj\n");
        WriteAscii(pdf, $"4 0 obj << /Length {content.Length} >> stream\n");
        pdf.Write(content, 0, content.Length);
        WriteAscii(pdf, "\nendstream endobj\n");
        string interpolation = interpolate ? " /Interpolate true" : string.Empty;
        WriteAscii(pdf, $"5 0 obj << /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter {filter}{interpolation} /Length {image.Length} >> stream\n");
        pdf.Write(image, 0, image.Length);
        WriteAscii(pdf, "\nendstream endobj\ntrailer << /Root 1 0 R /Size 6 >>\n%%EOF\n");
        return pdf.ToArray();
    }

    private static void WriteAscii(Stream stream, string value) {
        byte[] bytes = Encoding.ASCII.GetBytes(value);
        stream.Write(bytes, 0, bytes.Length);
    }
}
