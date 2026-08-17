using System.Text;
using System.Text.Json;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Web.Converter.Services;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Web.Converter.Tests;

public sealed class BrowserPdfImportTests {
    private readonly BrowserConversionService _service = new();

    [Theory]
    [InlineData("pdf-docx", ".docx")]
    [InlineData("pdf-xlsx", ".xlsx")]
    [InlineData("pdf-pptx", ".pptx")]
    [InlineData("pdf-html", ".html")]
    public void PdfImportRoutes_ProduceRealArtifactsAndReports(string routeId, string extension) {
        byte[] pdf = PdfDocument.Create(document => document.Content(content => content
            .H1("Quarterly report")
            .Paragraph(paragraph => paragraph.Text("Local PDF import evidence"))
            .Table([
                ["Metric", "Value"],
                ["Ready", "Yes"]
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
                break;
            case "pdf-html":
                string html = Encoding.UTF8.GetString(result.Bytes);
                Assert.Contains("Quarterly report", html, StringComparison.Ordinal);
                Assert.Equal(html, result.HtmlPreview);
                break;
        }
    }
}
