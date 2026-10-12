using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void SaveAsPdf_WorksheetCanvas_Uses_Workbook_Default_Font_For_Column_Units() {
        double defaultWidth = GetDefaultFontGeometryColumnWidth(10, 11);
        double largerDefaultWidth = GetDefaultFontGeometryColumnWidth(20, 11);
        double largerCellFontWidth = GetDefaultFontGeometryColumnWidth(10, 13);

        Assert.InRange(largerDefaultWidth / defaultWidth, 1.99, 2.01);
        Assert.InRange(Math.Abs(largerCellFontWidth - defaultWidth), 0, 0.02);
    }

    private double GetDefaultFontGeometryColumnWidth(double defaultSize, double cellSize) {
        string path = Path.Combine(_directoryWithFiles, $"ExcelPdfDefaultGeometry-{defaultSize}-{cellSize}.xlsx");
        using (ExcelDocument workbook = ExcelDocument.Create(path, "Geometry")) {
            ExcelSheet sheet = workbook.Sheets[0];
            sheet.Cell(1, 1, "ARef");
            sheet.Cell(1, 2, "BRef");
            sheet.SetColumnWidth(1, 12);
            sheet.SetColumnWidth(2, 12);
            sheet.SetRowHeight(1, 32);
            workbook.Save();
        }

        // Native column units use workbook font zero. Cells explicitly use a
        // different font, so changing their size must not resize the grid.
        using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
            Stylesheet styles = package.WorkbookPart!.WorkbookStylesPart!.Stylesheet;
            Font defaultFont = styles.Fonts!.Elements<Font>().First();
            defaultFont.GetFirstChild<FontName>()!.Val = "GeometryDefault";
            defaultFont.GetFirstChild<FontSize>()!.Val = defaultSize;
            styles.Fonts.Append(new Font(new FontName { Val = "Arial" }, new FontSize { Val = cellSize }));
            styles.Fonts.Count = (uint)styles.Fonts.Count();
            styles.CellFormats!.Elements<CellFormat>().First().FontId = 1U;
            styles.CellFormats!.Elements<CellFormat>().First().ApplyFont = true;
            styles.Save();
        }

        using ExcelDocument loaded = ExcelDocument.Load(path);
        var options = new ExcelToPdfOptions {
            WorksheetLayout = ExcelPdfWorksheetLayoutMode.WorksheetCanvas,
            IncludeSheetHeadings = false,
            HeaderRowCount = 0,
            PageSize = new PdfCore.PageSize(500, 200),
            Margins = PdfCore.PageMargins.Uniform(24),
            PdfOptions = new PdfCore.PdfOptions().RegisterNamedFontFamily(new PdfCore.PdfEmbeddedFontFamily(
                "GeometryDefault", OfficeIMO.TestAssets.PdfTestFontAssets.LoadBundledOpenTypeCffFont()))
        };
        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(loaded.ToPdfBytes(options)));
        var page = pdf.GetPage(1);
        return FindWordStartX(page, "BRef") - FindWordStartX(page, "ARef");
    }
}
