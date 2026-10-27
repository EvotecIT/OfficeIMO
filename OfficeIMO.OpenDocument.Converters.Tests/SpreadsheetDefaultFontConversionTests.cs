using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class SpreadsheetDefaultFontConversionTests {
    [Theory]
    [InlineData("120%", 14.4D)]
    [InlineData("80%", 9.6D)]
    [InlineData("14pt", 14D)]
    public void RelativeDefaultFontSizesUseTheOpenDocumentCascade(string lexical, double expectedPoints) {
        using var original = ExcelDocument.Create();
        original.AddWorksheet("Default font").CellValue(1, 1, "Visible cell");
        OdsDocument source = original.ToOpenDocument();
        source.Styles.GetOrCreateDefault(OdfStyleFamily.TableCell).FontSize = OdfLength.Parse(lexical);
        using var serialized = new MemoryStream();
        source.Save(serialized);
        serialized.Position = 0;
        OdsDocument loaded = OdsDocument.Load(serialized);

        using ExcelDocument imported = loaded.ToExcelDocument();

        Font font = imported.OpenXmlDocument.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.Fonts!.Elements<Font>().First();
        Assert.Equal(expectedPoints, font.FontSize!.Val!.Value, 6);
        Assert.Equal("Visible cell", imported.Inspect().Worksheets[0].Cells.Single().Value);
    }

    [Fact]
    public void WorkbookColumnFontSurvivesSavedOdsRoundTripIndependentlyOfCellFonts() {
        using var workbook = ExcelDocument.Create();
        var sheet = workbook.AddWorksheet("Metrics");
        sheet.CellValue(1, 1, "Visible cell");
        sheet.CellFontName(1, 1, "Cell Face");
        sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(24));
        Font font = workbook.OpenXmlDocument.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.Fonts!.Elements<Font>().First();
        font.FontName = new FontName { Val = "Column Face" };
        font.FontSize = new FontSize { Val = 14D };
        font.Bold = new Bold();
        font.Italic = new Italic();

        OdsDocument ods = workbook.ToOpenDocument();
        using var bytes = new MemoryStream();
        ods.Save(bytes);
        bytes.Position = 0;
        OdsDocument reopened = OdsDocument.Load(bytes);
        OdfStyle? defaults = reopened.Styles.FindDefault(OdfStyleFamily.TableCell);
        Assert.NotNull(defaults);
        Assert.Equal("Column Face", defaults!.FontFamily);
        Assert.Equal(14D, defaults.FontSize!.Value.ToPoints(), 6);
        Assert.False(reopened.Sheets[0].Cell(0, 0).Bold);
        Assert.False(reopened.Sheets[0].Cell(0, 0).Italic);
        using var roundTrip = reopened.ToExcelDocument();
        Font restored = roundTrip.OpenXmlDocument.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.Fonts!.Elements<Font>().First();
        Assert.Equal("Column Face", restored.FontName!.Val!.Value);
        Assert.Equal(14D, restored.FontSize!.Val!.Value);
        Assert.NotNull(restored.Bold);
        Assert.NotNull(restored.Italic);
        var cellStyle = roundTrip.Inspect().Worksheets[0].Cells.Single().Style!;
        Assert.Equal("Cell Face", cellStyle.FontName);
        Assert.False(cellStyle.Bold);
        Assert.False(cellStyle.Italic);
    }
}
