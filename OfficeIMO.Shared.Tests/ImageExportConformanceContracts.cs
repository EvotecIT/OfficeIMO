using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Visio;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class ImageExportConformanceContracts {
    [Fact]
    public void DocumentFamiliesShareFormatIdentityAndPreallocationSafety() {
        const long maximumPixels = 20_000L;
        var results = new List<OfficeImageExportResult>();

        using (var excelStream = new MemoryStream())
        using (ExcelDocument workbook = ExcelDocument.Create(excelStream)) {
            ExcelSheet sheet = workbook.AddWorksheet("Budget");
            sheet.CellValue(1, 1, "Excel");
            results.Add(sheet.Range("A1:D8").ExportImage(
                OfficeImageExportFormat.Webp,
                new ExcelImageExportOptions {
                    Scale = 50D,
                    MaximumRasterPixels = maximumPixels
                }));
        }

        HtmlConversionDocument html = HtmlConversionDocument.Parse(
            "<div style='width:200px;height:100px;background:#369;color:white'>HTML</div>");
        results.Add(html.ExportImage(
            OfficeImageExportFormat.Webp,
            new HtmlRenderOptions {
                Scale = 50D,
                ViewportWidth = 240D,
                MaximumRasterPixels = maximumPixels
            }));

        PdfReadDocument pdf = PdfReadDocument.Open(PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("PDF"))
            .ToBytes());
        results.Add(pdf.Pages[0].ExportImage(
            OfficeImageExportFormat.Webp,
            new PdfImageExportOptions {
                Scale = 50D,
                MaximumRasterPixels = maximumPixels
            }));

        using (var presentationStream = new MemoryStream())
        using (PowerPointPresentation presentation = PowerPointPresentation.Create(presentationStream)) {
            presentation.SlideSize.SetSizePoints(240D, 160D);
            PowerPointSlide slide = presentation.AddSlide();
            slide.BackgroundColor = "336699";
            results.Add(slide.ExportImage(
                OfficeImageExportFormat.Webp,
                new PowerPointImageExportOptions {
                    Scale = 50D,
                    MaximumRasterPixels = maximumPixels
                }));
        }

        using (var visioStream = new MemoryStream()) {
            VisioDocument document = VisioDocument.Create(visioStream);
            VisioPage page = document.AddPage("Budget").Size(2D, 1D);
            page.AddRectangle(1D, 0.5D, 1D, 0.5D, "Visio");
            results.Add(page.ExportImage(
                OfficeImageExportFormat.Webp,
                new VisioImageExportOptions {
                    Scale = 50D,
                    Supersampling = 1,
                    MaximumRasterPixels = maximumPixels
                }));
        }

        using (var wordStream = new MemoryStream())
        using (WordDocument document = WordDocument.Create(wordStream)) {
            document.AddParagraph("Word");
            results.Add(document.ExportImage(
                OfficeImageExportFormat.Webp,
                new WordImageExportOptions {
                    Scale = 50D,
                    MaximumRasterPixels = maximumPixels
                }));
        }

        Assert.Equal(6, results.Count);
        Assert.All(results, result => {
            Assert.Equal(OfficeImageExportFormat.Webp, result.Format);
            Assert.Equal("image/webp", result.MimeType);
            Assert.Equal(".webp", result.FileExtension);
            Assert.True((long)result.Width * result.Height <= maximumPixels);
            Assert.Contains(result.Diagnostics, diagnostic =>
                diagnostic.Code == OfficeImageExportDiagnosticCodes.RasterScaleReduced);
        });
    }

    [Fact]
    public void ResultRejectsMismatchedFormatBytesAtTheSharedBoundary() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));

        Assert.Throws<ArgumentException>(() =>
            new OfficeImageExportResult(
                OfficeImageExportFormat.Jpeg,
                1,
                1,
                png));
    }

    [Fact]
    public void ResultRejectsMismatchedDimensionsAtTheSharedBoundary() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(2, 1, OfficeColor.White));

        Assert.Throws<ArgumentException>(() =>
            new OfficeImageExportResult(
                OfficeImageExportFormat.Png,
                1,
                1,
                png));
    }
}
