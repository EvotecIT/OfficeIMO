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

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SaveAsPdf_WorksheetCanvas_Text_Substitution_Does_Not_Change_Column_Units(bool defaultFontAvailable) {
        string family = defaultFontAvailable ? "GeometryDefault" : "OfficeIMO Missing Column Font";
        double originalWidth = GetDefaultFontGeometryColumnWidth(10, 11, family);
        double substitutedWidth = GetDefaultFontGeometryColumnWidth(10, 11, family, substitute: true);

        Assert.InRange(Math.Abs(substitutedWidth - originalWidth), 0, 0.02);
    }

    [Fact]
    public void SaveAsPdf_WorksheetCanvas_Implicit_Columns_Use_Base_Width_And_Format_Padding() {
        double storedEight = GetDefaultFontGeometryColumnWidth(10, 11, defaultColumnWidth: 8);
        double baseEight = GetDefaultFontGeometryColumnWidth(10, 11, baseColumnWidth: 8);
        double baseTwelve = GetDefaultFontGeometryColumnWidth(10, 11, baseColumnWidth: 12);

        Assert.InRange(baseEight - storedEight, 3.73, 3.77);
        Assert.InRange(Math.Abs((baseTwelve - baseEight) - storedEight / 2), 0, 0.02);
    }

    [Fact]
    public void SaveAsPdf_WorksheetCanvas_Authored_Default_Width_Overrides_Base_Width() {
        double baseEight = GetDefaultFontGeometryColumnWidth(10, 11, baseColumnWidth: 8, defaultColumnWidth: 13);
        double baseTwelve = GetDefaultFontGeometryColumnWidth(10, 11, baseColumnWidth: 12, defaultColumnWidth: 13);

        Assert.InRange(Math.Abs(baseTwelve - baseEight), 0, 0.02);
    }

    private double GetDefaultFontGeometryColumnWidth(double defaultSize, double cellSize,
        string defaultFamily = "GeometryDefault", bool substitute = false, uint? baseColumnWidth = null,
        double? defaultColumnWidth = null) {
        string path = Path.Combine(_directoryWithFiles, $"ExcelPdfDefaultGeometry-{defaultSize}-{cellSize}-{defaultFamily}-{substitute}-{baseColumnWidth}-{defaultColumnWidth}.xlsx");
        using (ExcelDocument workbook = ExcelDocument.Create(path, new ExcelCreateOptions { FirstSheetName = "Geometry" })) {
            ExcelSheet sheet = workbook.Sheets[0];
            sheet.Cell(1, 1, "ARef");
            sheet.Cell(1, 2, "BRef");
            if (!baseColumnWidth.HasValue && !defaultColumnWidth.HasValue) {
                sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(12));
                sheet.SetColumnWidth(2, ExcelColumnWidth.Characters(12));
            }
            if (defaultColumnWidth.HasValue) sheet.SetDefaultColumnWidth(ExcelColumnWidth.Characters(defaultColumnWidth.Value));
            sheet.SetRowHeight(1, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(32)));
            workbook.Save();
        }

        // Native column units use workbook font zero. Cells explicitly use a
        // different font, so changing their size must not resize the grid.
        using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
            if (baseColumnWidth.HasValue) {
                Worksheet worksheet = package.WorkbookPart!.WorksheetParts.First().Worksheet;
                SheetFormatProperties? format = worksheet.GetFirstChild<SheetFormatProperties>();
                if (format == null) {
                    format = new SheetFormatProperties { DefaultRowHeight = 15 };
                    worksheet.AddChild(format, true);
                }
                format.BaseColumnWidth = baseColumnWidth.Value;
                worksheet.Save();
            }
            Stylesheet styles = package.WorkbookPart!.WorkbookStylesPart!.Stylesheet;
            Font defaultFont = styles.Fonts!.Elements<Font>().First();
            defaultFont.GetFirstChild<FontName>()!.Val = defaultFamily;
            defaultFont.GetFirstChild<FontSize>()!.Val = defaultSize;
            styles.Fonts.Append(new Font(new FontName { Val = "Arial" }, new FontSize { Val = cellSize }));
            styles.Fonts.Count = (uint)styles.Fonts.Count();
            styles.CellFormats!.Elements<CellFormat>().First().FontId = 1U;
            styles.CellFormats!.Elements<CellFormat>().First().ApplyFont = true;
            styles.Save();
        }

        using ExcelDocument loaded = ExcelDocument.Load(path);
        var options = new ExcelPdfExportOptions {
            WorksheetLayout = ExcelPdfWorksheetLayoutMode.WorksheetCanvas,
            IncludeSheetHeadings = false,
            HeaderRowCount = 0,
            PageSize = new PdfCore.PdfPageSize(500, 200),
            Margins = PdfCore.PdfPageMargins.Uniform(24),
            PdfOptions = new PdfCore.PdfCreateOptions().RegisterNamedFontFamily(new PdfCore.PdfEmbeddedFontFamily(
                "GeometryDefault", OfficeIMO.TestAssets.PdfTestFontAssets.LoadBundledOpenTypeCffFont()))
        };
        if (substitute) {
            options.PdfOptions.RegisterNamedFontFamily(new PdfCore.PdfEmbeddedFontFamily(
                "GeometryColumnSubstitute", OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFont(
                    Enumerable.Range('0', 10).ToArray())))
                .RegisterFontFamilySubstitution(defaultFamily, "GeometryColumnSubstitute",
                    PdfCore.PdfFontFamilySubstitutionImpact.LayoutSensitive);
        }
        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(loaded.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true })));
        var page = pdf.GetPage(1);
        return FindWordStartX(page, "BRef") - FindWordStartX(page, "ARef");
    }
}
