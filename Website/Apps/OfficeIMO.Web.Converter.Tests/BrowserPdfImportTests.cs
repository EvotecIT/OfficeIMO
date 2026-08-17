using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Web.Converter.Services;
using OfficeIMO.Word;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Web.Converter.Tests;

public sealed class BrowserPdfImportTests {
    private readonly BrowserConversionService _service = new();

    [Theory]
    [InlineData("pdf-docx", ".docx")]
    [InlineData("pdf-xlsx", ".xlsx")]
    [InlineData("pdf-pptx", ".pptx")]
    [InlineData("pdf-html", ".html")]
    [InlineData("pdf-png", ".png")]
    public void PdfImportRoutes_ProduceRealArtifactsAndReports(string routeId, string extension) {
        byte[] pdf = PdfDocument.Create(document => document.Content(content => content
            .H1("Quarterly report")
            .Paragraph(paragraph => paragraph.Text("Local PDF import evidence"))
            .Table([
                ["Metric", "Value"],
                ["Ready", "Yes"],
                ["Validated", "Yes"]
            ])))
            .ToBytes();
        var source = new SelectedDocument("report.pdf", ".pdf", "PDF", pdf.LongLength, pdf);

        ConversionResult result = _service.ConvertFile(ConversionRouteCatalog.Find(routeId), source, limitExcelRows: false);

        Assert.Equal("report" + extension, result.FileName);
        Assert.NotEmpty(result.Bytes);
        Assert.Equal(1, result.PageCount);
        Assert.NotNull(result.CompanionReport);
        Assert.Equal("report.officeimo-report.json", result.CompanionReport!.FileName);
        using JsonDocument report = JsonDocument.Parse(result.CompanionReport!.Bytes);
        Assert.Equal(routeId, report.RootElement.GetProperty("route").GetString());
        Assert.True(report.RootElement.GetProperty("browserLocal").GetBoolean());
        VerifyArtifact(routeId, result);
    }

    private static void VerifyArtifact(string routeId, ConversionResult result) {
        switch (routeId) {
            case "pdf-docx":
                using (WordDocument word = WordDocument.Load(new MemoryStream(result.Bytes))) {
                    Assert.Contains("Quarterly report", word.Paragraphs.Select(static paragraph => paragraph.Text));
                }
                break;
            case "pdf-xlsx":
                using (ExcelDocument excel = ExcelDocument.Load(new MemoryStream(result.Bytes))) {
                    Assert.NotEmpty(excel.Sheets);
                }
                break;
            case "pdf-pptx":
                using (PowerPointPresentation presentation = PowerPointPresentation.Load(new MemoryStream(result.Bytes))) {
                    Assert.NotEmpty(presentation.Slides);
                }
                using (PresentationDocument package = PresentationDocument.Open(new MemoryStream(result.Bytes), false)) {
                    SlidePart slide = Assert.Single(package.PresentationPart!.SlideParts);
                    Assert.Empty(slide.Slide!.Descendants<DocumentFormat.OpenXml.Presentation.Picture>());
                    Assert.Single(slide.Slide.Descendants<A.Table>());
                    Assert.Contains(slide.Slide.Descendants<A.Text>(), text => text.Text == "Quarterly report");
                }
                Assert.Equal("Degraded", result.FidelityStatus);
                using (JsonDocument report = JsonDocument.Parse(result.CompanionReport!.Bytes)) {
                    Assert.Equal("editable-content-slides", report.RootElement.GetProperty("projection").GetString());
                }
                break;
            case "pdf-html":
                string html = Encoding.UTF8.GetString(result.Bytes);
                Assert.Contains("Quarterly report", html, StringComparison.Ordinal);
                Assert.Contains("class=\"pdf-page\"", html, StringComparison.Ordinal);
                Assert.Equal(html, result.HtmlPreview);
                Assert.Contains("positioned PDF review reconstruction", result.ProvenanceSummary, StringComparison.Ordinal);
                Assert.DoesNotContain("semantic PDF reconstruction", result.ProvenanceSummary, StringComparison.Ordinal);
                using (JsonDocument report = JsonDocument.Parse(result.CompanionReport!.Bytes)) {
                    Assert.Equal("pdf-html-positioned-review", report.RootElement.GetProperty("projection").GetString());
                }
                break;
            case "pdf-png":
                Assert.Equal("image/png", result.ContentType);
                Assert.Contains("visual PDF page rendering", result.ProvenanceSummary, StringComparison.Ordinal);
                Assert.DoesNotContain("logical PDF import", result.ProvenanceSummary, StringComparison.Ordinal);
                Assert.True(OfficeRasterImageDecoder.TryDecode(result.Bytes, out OfficeRasterImage? raster));
                Assert.NotNull(raster);
                Assert.True(raster!.Width > 0);
                Assert.True(raster.Height > 0);
                using (JsonDocument report = JsonDocument.Parse(result.CompanionReport!.Bytes)) {
                    Assert.Equal("visual-page-images", report.RootElement.GetProperty("projection").GetString());
                }
                break;
        }
    }

    [Fact]
    public void PdfToPng_MultiplePages_ReturnsNumberedPngArchive() {
        PdfDocument document = PdfDocument.Create(compose => {
            compose.Page(page => page.Content(content => content.Column(column =>
                column.Item().Paragraph(paragraph => paragraph.Text("First page")))));
            compose.Page(page => page.Content(content => content.Column(column =>
                column.Item().Paragraph(paragraph => paragraph.Text("Second page")))));
        });
        byte[] pdf = document.ToBytes();
        var source = new SelectedDocument("pages.pdf", ".pdf", "PDF", pdf.LongLength, pdf);

        ConversionResult result = _service.ConvertFile(
            ConversionRouteCatalog.Find("pdf-png"),
            source,
            limitExcelRows: false);

        Assert.Equal("pages.zip", result.FileName);
        Assert.Equal("application/zip", result.ContentType);
        Assert.Equal(2, result.PageCount);
        Assert.Contains("visual PDF page rendering", result.ProvenanceSummary, StringComparison.Ordinal);
        using var archive = new ZipArchive(new MemoryStream(result.Bytes), ZipArchiveMode.Read);
        Assert.Equal(["pages.page-001.png", "pages.page-002.png"], archive.Entries.Select(static entry => entry.FullName));
        foreach (ZipArchiveEntry entry in archive.Entries) {
            using Stream stream = entry.Open();
            using var image = new MemoryStream();
            stream.CopyTo(image);
            Assert.True(OfficeRasterImageDecoder.TryDecode(image.ToArray(), out OfficeRasterImage? raster));
            Assert.NotNull(raster);
        }
        using JsonDocument report = JsonDocument.Parse(result.CompanionReport!.Bytes);
        Assert.Equal("pages.zip", report.RootElement.GetProperty("output").GetProperty("fileName").GetString());
    }

    [Theory]
    [InlineData(PdfPowerPointImportMode.VisualPages, "Visual", "visual-page-slides", true, false)]
    [InlineData(PdfPowerPointImportMode.HybridVisualAndEditableTables, "Hybrid", "hybrid-visual-table-slides", true, true)]
    [InlineData(PdfPowerPointImportMode.EditableTables, "Partial", "editable-table-slides", false, true)]
    [InlineData(PdfPowerPointImportMode.EditableContent, "Degraded", "editable-content-slides", false, true)]
    public void PdfToPowerPoint_ExposesHonestProjectionModes(
        PdfPowerPointImportMode mode,
        string fidelity,
        string projection,
        bool expectsPagePicture,
        bool expectsTable) {
        byte[] pdf = PdfDocument.Create(document => document.Content(content => content
            .H1("Mode contract")
            .Table([
                ["Metric", "Value"],
                ["Ready", "Yes"],
                ["Validated", "Yes"]
            ])))
            .ToBytes();
        var source = new SelectedDocument("modes.pdf", ".pdf", "PDF", pdf.LongLength, pdf);

        ConversionResult result = _service.ConvertFile(
            ConversionRouteCatalog.Find("pdf-pptx"),
            source,
            limitExcelRows: false,
            pdfPowerPointMode: mode);

        Assert.True(
            string.Equals(fidelity, result.FidelityStatus, StringComparison.Ordinal),
            string.Join(Environment.NewLine, result.StructuredWarnings.Select(static warning => $"{warning.Code}: {warning.Message}")));
        using JsonDocument report = JsonDocument.Parse(result.CompanionReport!.Bytes);
        Assert.Equal(projection, report.RootElement.GetProperty("projection").GetString());
        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(result.Bytes), false);
        SlidePart slide = Assert.Single(package.PresentationPart!.SlideParts);
        Assert.Equal(expectsPagePicture, slide.Slide!.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Any());
        Assert.Equal(expectsTable, slide.Slide.Descendants<A.Table>().Any());
        if (mode == PdfPowerPointImportMode.VisualPages) {
            Assert.Contains(result.StructuredWarnings, warning => warning.Code == "PdfVisualPageSlidesNotEditable");
            Assert.Contains("not editable", result.Text!, StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public void PdfToPowerPoint_VisualMode_PreservesShowcaseBase14PositionedText() {
        string sourcePath = Path.Combine(AppContext.BaseDirectory, "samples", "showcase-dashboard.pdf");
        byte[] pdf = File.ReadAllBytes(sourcePath);
        var source = new SelectedDocument("OfficeIMO-Showcase.pdf", ".pdf", "PDF", pdf.LongLength, pdf);

        PdfReadDocument readDocument = PdfReadDocument.Open(pdf);
        OfficeDrawing drawing = readDocument.Pages[0].ToDrawing();
        OfficeDrawingText[] sourceText = EnumerateText(drawing).ToArray();
        Assert.Contains(sourceText, static text => text.Text == "Quarterly");
        Assert.Contains(sourceText, static text => text.Text == "Operations");
        Assert.Contains(sourceText, static text => text.Text == "Dashboard");
        Assert.Contains(sourceText, static text => text.Text == "OfficeIMO.Pdf");
        Assert.All(
            sourceText.Where(static text => Math.Abs(text.RotationDegrees) <= 0.0001D),
            static text => {
                Assert.Equal(OfficeTextOverflowBehavior.Clip, text.OverflowBehavior);
                Assert.True(text.TextAdvanceWidth.HasValue && text.TextAdvanceWidth.Value > 0D);
            });

        ConversionResult result = _service.ConvertFile(
            ConversionRouteCatalog.Find("pdf-pptx"),
            source,
            limitExcelRows: false,
            pdfPowerPointMode: PdfPowerPointImportMode.VisualPages);

        Assert.Equal("Visual", result.FidelityStatus);
        Assert.Contains(result.StructuredWarnings, warning =>
            warning.Code.Contains("font", StringComparison.OrdinalIgnoreCase));
        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(result.Bytes), false);
        SlidePart slide = Assert.Single(package.PresentationPart!.SlideParts);
        ImagePart image = Assert.Single(slide.ImageParts);
        using Stream imageStream = image.GetStream(FileMode.Open, FileAccess.Read);
        using var imageBuffer = new MemoryStream();
        imageStream.CopyTo(imageBuffer);
        byte[] renderedPage = imageBuffer.ToArray();
        Assert.True(OfficeIMO.Drawing.OfficeRasterImageDecoder.TryDecode(renderedPage, out OfficeIMO.Drawing.OfficeRasterImage? raster));
        Assert.NotNull(raster);
        Assert.Equal(1684, raster!.Width);
        Assert.Equal(1190, raster.Height);
        Assert.Equal(
            "86c8cc7de11d99412b02da94067779a58c19fca2e38e54456dc0bdd1c5d39abb",
            Convert.ToHexString(SHA256.HashData(renderedPage)).ToLowerInvariant());
    }

    [Fact]
    public void PdfToPng_MatchesTheShowcaseVisualPowerPointPagePixels() {
        string sourcePath = Path.Combine(AppContext.BaseDirectory, "samples", "showcase-dashboard.pdf");
        byte[] pdf = File.ReadAllBytes(sourcePath);
        var source = new SelectedDocument("OfficeIMO-Showcase.pdf", ".pdf", "PDF", pdf.LongLength, pdf);

        ConversionResult result = _service.ConvertFile(
            ConversionRouteCatalog.Find("pdf-png"),
            source,
            limitExcelRows: false);

        Assert.Equal("OfficeIMO-Showcase.png", result.FileName);
        Assert.Equal("image/png", result.ContentType);
        Assert.True(OfficeRasterImageDecoder.TryDecode(result.Bytes, out OfficeRasterImage? raster));
        Assert.NotNull(raster);
        Assert.Equal(1684, raster!.Width);
        Assert.Equal(1190, raster.Height);

        ConversionResult visualSlides = _service.ConvertFile(
            ConversionRouteCatalog.Find("pdf-pptx"),
            source,
            limitExcelRows: false,
            pdfPowerPointMode: PdfPowerPointImportMode.VisualPages);
        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(visualSlides.Bytes), false);
        ImagePart image = Assert.Single(package.PresentationPart!.SlideParts.Single().ImageParts);
        using Stream imageStream = image.GetStream(FileMode.Open, FileAccess.Read);
        using var imageBuffer = new MemoryStream();
        imageStream.CopyTo(imageBuffer);
        Assert.True(OfficeRasterImageDecoder.TryDecode(imageBuffer.ToArray(), out OfficeRasterImage? slideRaster));
        Assert.NotNull(slideRaster);
        Assert.Equal(slideRaster!.Width, raster.Width);
        Assert.Equal(slideRaster.Height, raster.Height);
        Assert.Equal(
            Convert.ToHexString(SHA256.HashData(slideRaster.GetPixels())),
            Convert.ToHexString(SHA256.HashData(raster.GetPixels())));
    }

    [Fact]
    public void PdfToPowerPoint_VisualMode_RendersEveryBase14TextFamilyWithPinnedBrowserFallback() {
        byte[] pdf = PdfDocument.Create(document => document.Content(content => content
            .Paragraph(p => p.Runs([PdfTextRun.Normal("Helvetica regular", font: PdfStandardFont.Helvetica)]))
            .Paragraph(p => p.Runs([PdfTextRun.Normal("Helvetica bold italic", font: PdfStandardFont.HelveticaBoldOblique)]))
            .Paragraph(p => p.Runs([PdfTextRun.Normal("Times regular", font: PdfStandardFont.TimesRoman)]))
            .Paragraph(p => p.Runs([PdfTextRun.Normal("Times bold italic", font: PdfStandardFont.TimesBoldItalic)]))
            .Paragraph(p => p.Runs([PdfTextRun.Normal("Courier regular", font: PdfStandardFont.Courier)]))
            .Paragraph(p => p.Runs([PdfTextRun.Normal("Courier bold italic", font: PdfStandardFont.CourierBoldOblique)]))), new PdfOptions {
                PageWidth = 500,
                PageHeight = 600,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 14
            })
            .ToBytes();
        var source = new SelectedDocument("base14.pdf", ".pdf", "PDF", pdf.LongLength, pdf);

        ConversionResult result = _service.ConvertFile(
            ConversionRouteCatalog.Find("pdf-pptx"),
            source,
            limitExcelRows: false,
            pdfPowerPointMode: PdfPowerPointImportMode.VisualPages);

        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(result.Bytes), false);
        ImagePart image = Assert.Single(package.PresentationPart!.SlideParts.Single().ImageParts);
        using Stream imageStream = image.GetStream(FileMode.Open, FileAccess.Read);
        using var imageBuffer = new MemoryStream();
        imageStream.CopyTo(imageBuffer);
        byte[] renderedPage = imageBuffer.ToArray();
        Assert.True(OfficeIMO.Drawing.OfficeRasterImageDecoder.TryDecode(renderedPage, out OfficeIMO.Drawing.OfficeRasterImage? raster));
        Assert.NotNull(raster);
        Assert.Equal(1000, raster!.Width);
        Assert.Equal(1200, raster.Height);
        Assert.Equal(
            "b545034ba0c75db86e6fbff8333115e1ca6cd42fbe86e70fb1ac1da0324f00df",
            Convert.ToHexString(SHA256.HashData(renderedPage)).ToLowerInvariant());
    }

    private static IEnumerable<OfficeDrawingText> EnumerateText(OfficeDrawing drawing) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            if (element is OfficeDrawingText text) {
                yield return text;
            } else if (element is OfficeDrawingGroup group) {
                foreach (OfficeDrawingText child in EnumerateText(group.Drawing)) {
                    yield return child;
                }
            }
        }
    }
}
