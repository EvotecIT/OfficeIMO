using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Drawing;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Excel {

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Rejects_Invalid_Options() {
        Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelPdfExportOptions { HeaderRowCount = -1 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelPdfExportOptions { MaxRowsPerSheet = 0 });
    }

    [Fact]
    public void SaveAsPdf_ExcelShapingOnlyProfileAllowsWorkbookFontDiscoveryUntilCallerConfiguresFonts() {
        var options = new ExcelPdfExportOptions()
            .UseRenderingProfile(new OfficeRenderingProfile("shaping-only"));

        Assert.False(options.HasExplicitPdfFontConfiguration);

        options.PdfOptions!.DefaultFontSize = 13;

        Assert.True(options.HasExplicitPdfFontConfiguration);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Reports_Unsupported_Export_Features() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfUnsupportedFeatureWarnings.xlsx");

        var options = new ExcelPdfExportOptions {
            IncludeSheetHeadings = false,
            HeaderRowCount = 1,
            MaxRowsPerSheet = 2,
            PageSize = new PdfCore.PdfPageSize(460, 320),
            Margins = PdfCore.PdfPageMargins.Uniform(24)
        };

        byte[] bytes;
        PdfCore.PdfDocumentConversionResult result;
        PdfCore.PdfSaveResult saveResult;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, new ExcelCreateOptions { FirstSheetName = "Warnings" })) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Name");
            sheet.Cell(1, 2, "Value");
            sheet.Cell(2, 1, "Alpha");
            sheet.Cell(2, 2, 10);
            sheet.Cell(3, 1, "Beta");
            sheet.Cell(3, 2, 20);
            sheet.SetHeaderFooter(
                headerCenter: "&U&\"Arial,Bold\"&14&KFF0000Styled &D &T &A",
                footerRight: "Page &P of &N");
            sheet.AddBubbleChartFromRanges(new[] { new ExcelChartSeriesRange("Values", "B2:B3", "B2:B3", "B2:B3") },
                row: 1, column: 4, widthPixels: 320, heightPixels: 180, title: "Unsupported Bubble Chart");

            ExcelChart chart = Assert.Single(sheet.Charts);
            // Truncated producer data must remain an unreadable-chart diagnostic.
            document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.ChartParts.Single()
                .ChartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.BubbleSize>().Single().Remove();
            Assert.False(chart.TryGetSnapshot(out _));

            document.Save();

            result = document.ConvertToPdfWithEvidence(options);
            bytes = result.ToBytes();
            using var output = new MemoryStream();
            saveResult = document.ConvertToPdfWithEvidence(options).Save(output);
        }

        using (PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes))) {
            string text = pdf.GetPage(1).Text;
            Assert.Contains("Styled", text);
            Assert.Contains(DateTime.Now.ToString("d", CultureInfo.CurrentCulture), text);
            Assert.Contains("Warnings", text);
            Assert.Contains("Page 1 of", text);
            Assert.Contains("Alpha", text);
            Assert.DoesNotContain("Beta", text);
            Assert.DoesNotContain("Unsupported Bubble Chart", text);
        }

        Assert.Contains(result.Warnings, warning => warning.Source == "Warnings" && warning.Code == "worksheet-header-footer-formatting");
        Assert.Contains(result.Warnings, warning => warning.Source == "Warnings" && warning.Code == "worksheet-rows");
        Assert.Contains(result.Warnings, warning => warning.Source == "Warnings" && warning.Code == "worksheet-chart" && warning.Message.Contains("could not be read", StringComparison.Ordinal));
        Assert.True(saveResult.Succeeded);
        Assert.Contains(saveResult.Warnings, warning => warning.Source == "Warnings" && warning.Code == "worksheet-header-footer-formatting");
        Assert.Contains(saveResult.Warnings, warning => warning.Source == "Warnings" && warning.Code == "worksheet-rows");
        Assert.Contains(saveResult.Warnings, warning => warning.Source == "Warnings" && warning.Code == "worksheet-chart");
    }
}
