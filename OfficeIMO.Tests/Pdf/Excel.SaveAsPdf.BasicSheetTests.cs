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
    public void SaveAsPdf_ExcelWorkbook_Exports_Worksheet_UsedRange_To_Table() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfWorkbook.xlsx");
        string pdfPath = Path.Combine(_directoryWithFiles, "ExcelPdfWorkbook.pdf");

        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Sales")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Product");
            sheet.Cell(1, 2, "Amount");
            sheet.Cell(2, 1, "Licenses");
            sheet.Cell(2, 2, 1250.5);
            sheet.Cell(3, 1, "Support");
            sheet.Cell(3, 2, 250);
            document.Save(false);

            document.SaveAsPdf(pdfPath);
        }

        Assert.True(File.Exists(pdfPath));
        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Sales", text);
        Assert.Contains("Product", text);
        Assert.Contains("Licenses", text);
        Assert.Contains("1250.5", text);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Respects_Selected_Sheets() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfSelectedSheets.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath)) {
            ExcelSheet summary = document.AddWorkSheet("Summary");
            summary.Cell(1, 1, "Metric");
            summary.Cell(2, 1, "SelectedValue");
            summary.SetHeaderFooter(headerCenter: "Selected Header &A");
            ExcelSheet internalSheet = document.AddWorkSheet("Internal");
            internalSheet.Cell(1, 1, "HiddenValue");
            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                SheetNames = new[] { "summary" }
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Summary", text);
        Assert.Contains("Selected Header Summary", text);
        Assert.Contains("SelectedValue", text);
        Assert.DoesNotContain("HiddenValue", text);
    }


    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Omits_Hidden_Sheets_By_Default() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfHiddenSheets.xlsx");

        byte[] visibleBytes;
        byte[] explicitHiddenBytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath)) {
            ExcelSheet visible = document.AddWorkSheet("Visible");
            visible.Cell(1, 1, "VisibleSheetValue");
            ExcelSheet hidden = document.AddWorkSheet("Hidden");
            hidden.Cell(1, 1, "HiddenSheetValue");
            hidden.SetHidden(true);
            Assert.True(hidden.Hidden);
            document.Save(false);

            visibleBytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 0
            });

            explicitHiddenBytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                SheetNames = new[] { "Hidden" },
                IncludeSheetHeadings = false,
                HeaderRowCount = 0
            });
        }

        using PdfPigDocument visiblePdf = PdfPigDocument.Open(new MemoryStream(visibleBytes));
        string visibleText = visiblePdf.GetPage(1).Text;
        Assert.Contains("VisibleSheetValue", visibleText);
        Assert.DoesNotContain("HiddenSheetValue", visibleText);

        using PdfPigDocument explicitHiddenPdf = PdfPigDocument.Open(new MemoryStream(explicitHiddenBytes));
        string hiddenText = explicitHiddenPdf.GetPage(1).Text;
        Assert.Contains("HiddenSheetValue", hiddenText);
    }

}
