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
    public void SaveAsPdf_WorksheetCanvas_Uses_Normal_Style_Font_For_Column_Units() {
        double normalWidth = GetNormalFontGeometryColumnWidth(10, 11);
        double largerNormalWidth = GetNormalFontGeometryColumnWidth(20, 11);
        double largerCellFontWidth = GetNormalFontGeometryColumnWidth(10, 13);

        Assert.InRange(largerNormalWidth / normalWidth, 1.99, 2.01);
        Assert.InRange(Math.Abs(largerCellFontWidth - normalWidth), 0, 0.02);
    }

    private double GetNormalFontGeometryColumnWidth(double normalSize, double cellSize) {
        string path = Path.Combine(_directoryWithFiles, $"ExcelPdfNormalGeometry-{normalSize}-{cellSize}.xlsx");
        using (ExcelDocument workbook = ExcelDocument.Create(path, "Geometry")) {
            ExcelSheet sheet = workbook.Sheets[0];
            sheet.Cell(1, 1, "ARef");
            sheet.Cell(1, 2, "BRef");
            sheet.SetColumnWidth(1, 12);
            sheet.SetColumnWidth(2, 12);
            sheet.SetRowHeight(1, 32);
            workbook.Save();
        }

        // The Normal style deliberately uses a nonzero font and style-format
        // index. Cell font zero remains explicit and independent of its units.
        using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
            Stylesheet styles = package.WorkbookPart!.WorkbookStylesPart!.Stylesheet;
            Font cellFont = styles.Fonts!.Elements<Font>().First();
            cellFont.GetFirstChild<FontName>()!.Val = "Arial";
            cellFont.GetFirstChild<FontSize>()!.Val = cellSize;
            styles.Fonts.Append(new Font(new FontName { Val = "GeometryNormal" }, new FontSize { Val = normalSize }));
            styles.Fonts.Count = (uint)styles.Fonts.Count();
            styles.CellStyleFormats!.Append(new CellFormat { FontId = 1U, FillId = 0U, BorderId = 0U, NumberFormatId = 0U });
            styles.CellStyleFormats.Count = (uint)styles.CellStyleFormats.Count();
            styles.CellStyles!.Elements<CellStyle>().Single(style => style.BuiltinId?.Value == 0U).FormatId = 1U;
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
                "GeometryNormal", OfficeIMO.TestAssets.PdfTestFontAssets.LoadBundledOpenTypeCffFont()))
        };
        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(loaded.ToPdfBytes(options)));
        var page = pdf.GetPage(1);
        return FindWordStartX(page, "BRef") - FindWordStartX(page, "ARef");
    }
}
