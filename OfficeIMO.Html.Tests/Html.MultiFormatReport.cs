using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlMultiFormatReport {
    [Fact]
    public void OnePreparedReportProducesSearchablePdfAndEditableOfficeArtifacts() {
        string sourcePath = Path.Combine(AppContext.BaseDirectory, "Reports", "service-review.html");
        HtmlConversionDocument report = HtmlConversionDocument.Load(sourcePath);
        string originalHtml = report.SourceHtml;

        byte[] pdfBytes = report.ToPdfBytes();
        using (var pdf = UglyToad.PdfPig.PdfDocument.Open(pdfBytes)) {
            Assert.Equal(1, pdf.NumberOfPages);
            string text = string.Join(" ", pdf.GetPages().SelectMany(page => page.GetWords()).Select(word => word.Text));
            Assert.Contains("Monthly service review", text, StringComparison.Ordinal);
            Assert.Contains("46.75", text, StringComparison.Ordinal);
            Assert.Contains("00130", text, StringComparison.Ordinal);
        }

        HtmlToWordResult wordResult = report.ToWordDocumentResult();
        using WordDocument word = wordResult.Value;
        wordResult.RequireValue();
        using MemoryStream wordArtifact = word.ToStream();
        using (WordprocessingDocument package = WordprocessingDocument.Open(wordArtifact, false)) {
            var errors = new OpenXmlValidator().Validate(package).ToArray();
            Assert.True(errors.Length == 0, string.Join(Environment.NewLine,
                errors.Select(error => error.Description + " at " + error.Path?.XPath + ": " + error.Node?.OuterXml)));
            var body = package.MainDocumentPart!.Document.Body!;
            Assert.Contains("Monthly service review", body.InnerText, StringComparison.Ordinal);
            Assert.Equal(2, body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>().Count());
            Assert.Contains(body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>(), text => text.Text == "00130");
            Assert.NotEmpty(package.MainDocumentPart.HyperlinkRelationships);
            var header = Assert.Single(body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>(),
                run => run.InnerText == "Workstream");
            Assert.Equal("FFFFFF", header.RunProperties?.Color?.Val?.Value);
            Assert.NotNull(header.RunProperties?.Bold);
        }

        HtmlToExcelResult excelResult = report.ToExcelDocumentResult(new HtmlToExcelOptions {
            Mode = HtmlImportMode.Generic,
            ImportTypedCellValues = true
        });
        using ExcelDocument excel = excelResult.Value;
        excelResult.RequireValue();
        using MemoryStream excelArtifact = excel.ToStream();
        byte[] excelBytes = excelArtifact.ToArray();
        using (SpreadsheetDocument package = SpreadsheetDocument.Open(new MemoryStream(excelBytes), false)) {
            Assert.Empty(new OpenXmlValidator().Validate(package));
        }
        using ExcelDocument reopened = ExcelDocument.Load(new MemoryStream(excelBytes));
        ExcelSheet usage = Assert.Single(reopened.Sheets, sheet => sheet.Name == "Service usage");
        Assert.Equal(46.75D, Enumerable.Range(2, 4).Sum(row => usage.CellAt(row, 3).GetValue<double>()));
        Assert.Equal(46.75D, usage.CellAt(6, 3).GetValue<double>());
        Assert.Equal("00130", usage.CellAt(5, 2).GetValue<string>());
        Assert.False(usage.CellAt(5, 4).GetValue<bool>());
        Assert.Equal(new[] { "A2:A3", "A4:A5", "A6:B6", "D6:E6" },
            usage.GetMergedRanges().Select(range => range.A1Range).OrderBy(range => range).ToArray());
        ExcelSheet actions = Assert.Single(reopened.Sheets, sheet => sheet.Name == "Next actions");
        Assert.True(actions.TryGetCellValueSnapshot(2, 3, out ExcelCellValueSnapshot? due));
        Assert.Equal(new DateTime(2026, 9, 8), due!.DateTimeValue);
        Assert.Contains(reopened.Sheets, sheet => sheet.Name == "Imported");
        Assert.Equal(0, excelResult.Formulas);
        Assert.Equal(originalHtml, report.SourceHtml);
    }
}
