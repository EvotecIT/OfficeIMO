using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
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
        Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelPdfSaveOptions { HeaderRowCount = -1 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelPdfSaveOptions { MaxRowsPerSheet = 0 });
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Reports_Unsupported_Export_Features() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfUnsupportedFeatureWarnings.xlsx");

        var options = new ExcelPdfSaveOptions {
            IncludeSheetHeadings = false,
            HeaderRowCount = 1,
            MaxRowsPerSheet = 2,
            PageSize = new PdfCore.PageSize(460, 320),
            Margins = PdfCore.PageMargins.Uniform(24)
        };

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Warnings")) {
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
            sheet.AddChartFromRange("A1:B3", row: 1, column: 4, widthPixels: 320, heightPixels: 180, type: ExcelChartType.Surface, title: "Unsupported Surface Chart");

            ExcelChart chart = Assert.Single(sheet.Charts);
            Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
            Assert.Equal(ExcelChartType.Surface, snapshot.ChartType);

            document.Save(false);

            bytes = document.SaveAsPdf(options);
        }

        using (PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes))) {
            string text = pdf.GetPage(1).Text;
            Assert.Contains("Styled", text);
            Assert.Contains(DateTime.Now.ToString("d", CultureInfo.CurrentCulture), text);
            Assert.Contains("Warnings", text);
            Assert.Contains("Page 1 of", text);
            Assert.Contains("Alpha", text);
            Assert.DoesNotContain("Beta", text);
            Assert.DoesNotContain("Unsupported Surface Chart", text);
        }

        Assert.Contains(options.Warnings, warning => warning.SheetName == "Warnings" && warning.Feature == "WorksheetHeaderFooterFormatting");
        Assert.Contains(options.Warnings, warning => warning.SheetName == "Warnings" && warning.Feature == "WorksheetRows");
        Assert.Contains(options.Warnings, warning => warning.SheetName == "Warnings" && warning.Feature == "WorksheetChart" && warning.Message.Contains("Surface", StringComparison.Ordinal));
        Assert.All(options.Warnings, warning => Assert.Contains("Warnings", warning.ToString(), StringComparison.Ordinal));
    }
}
