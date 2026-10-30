using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void SaveAsPdf_UsesExcelNormalMarginsAndAlignsRunningLinesWithTheSheet() {
        string path = Path.Combine(_directoryWithFiles, "ExcelNormalMargins.xlsx");
        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(path, new ExcelCreateOptions { FirstSheetName = "Margins" })) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "BodyText");
            sheet.SetHeaderFooter(headerLeft: "HeaderLeft", headerRight: "HeaderRight", footerLeft: "FooterLeft");
            document.Save();
            bytes = document.ToPdf(new ExcelPdfExportOptions { IncludeSheetHeadings = false })
                .ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        double pageHeight = page.Height;
        double left = 0.7D * 72D;

        // The body starts at Excel's 0.7 inch left margin plus the cell's text padding.
        double bodyX = FindWordStartX(page, "BodyText");
        Assert.InRange(bodyX, left, left + 4D);

        // The running lines share the sheet's left and right edges.
        Assert.Equal(left, FindWordStartX(page, "HeaderLeft"), 1);
        Assert.Equal(left, FindWordStartX(page, "FooterLeft"), 1);
        double headerRightEnd = page.Letters.Where(letter => letter.Value == "t").Max(letter => letter.EndBaseLine.X);
        Assert.Equal(page.Width - left, headerRightEnd, 0);

        // The header sits just below Excel's 0.3 inch header margin and above the 0.75 inch body margin.
        double headerBaselineFromTop = pageHeight - FindWordStartY(page, "HeaderLeft");
        Assert.InRange(headerBaselineFromTop, 0.3D * 72D + 4D, 0.75D * 72D);
        double footerBaselineFromBottom = FindWordStartY(page, "FooterLeft");
        Assert.InRange(footerBaselineFromBottom, 0.3D * 72D, 0.3D * 72D + 4D);
    }

    [Fact]
    public void SaveAsPdf_ClipsASingleColumnWiderThanThePageInsteadOfContinuingIt() {
        string path = Path.Combine(_directoryWithFiles, "ExcelWideSingleColumn.xlsx");
        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(path, new ExcelCreateOptions { FirstSheetName = "Wide" })) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.SetColumnWidth(1, 120);
            sheet.Cell(1, 1, "Wide column text that ends past the right margin of a portrait letter page in this export test case");
            document.Save();
            bytes = document.ToPdf(new ExcelPdfExportOptions { IncludeSheetHeadings = false })
                .ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(1, pdf.NumberOfPages);
    }

    [Fact]
    public void AddTable_RightAlignsHeadersOverCalculatedNumberColumns() {
        string path = Path.Combine(_directoryWithFiles, "ExcelFormulaColumnHeaders.xlsx");
        using ExcelDocument document = ExcelDocument.Create(path, new ExcelCreateOptions { FirstSheetName = "Data" });
        ExcelSheet sheet = document.Sheets[0];
        sheet.Cell(1, 1, "Item"); sheet.Cell(1, 2, "Cost"); sheet.Cell(1, 3, "Margin"); sheet.Cell(1, 4, "Label");
        for (int row = 2; row <= 4; row++) {
            sheet.Cell(row, 1, "Row " + row);
            sheet.Cell(row, 2, row * 10D, numberFormat: "#,##0");
            sheet.CellFormula(row, 3, $"B{row}*2");
            sheet.FormatCell(row, 3, "#,##0");
            sheet.CellFormula(row, 4, $"A{row}&\"!\"");
        }

        sheet.AddTable("A1:D4", hasHeader: true, name: "Costs", style: ExcelTableStyle.TableStyleMedium2);

        Assert.Equal("right", sheet.GetCellStyle(1, 2).HorizontalAlignment);
        Assert.Equal("right", sheet.GetCellStyle(1, 3).HorizontalAlignment);
        Assert.Null(sheet.GetCellStyle(1, 4).HorizontalAlignment);
        Assert.Null(sheet.GetCellStyle(1, 1).HorizontalAlignment);
    }
}
