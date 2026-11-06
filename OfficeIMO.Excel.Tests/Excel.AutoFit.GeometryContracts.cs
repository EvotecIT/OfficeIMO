using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public class ExcelAutoFitGeometryContractTests {
    [Fact]
    public void AutoFitStackedTextMeasuresGraphemesInsteadOfUtf16Units() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Stacked");
        sheet.Cell(1, 1).SetValue("A😀B").SetFontFamily("Calibri").SetTextRotation(255);
        sheet.AutoFitColumns();
        sheet.AutoFitRows();
        using var package = SpreadsheetDocument.Open(new MemoryStream(document.ToBytes()), false);
        var worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet;
        var measurer = OfficeTextMeasurer.Create(new OfficeFontInfo("Calibri", 11));
        var style = measurer.DefaultStyle;
        double lineHeight = Math.Ceiling(measurer.MeasureLineHeightOrDefault("Xg", style, 0) + 2);
        double expectedHeight = Math.Max(15d, lineHeight * .75) * 3;
        double glyphWidth = new[] { "A", "😀", "B" }.Max(text => measurer.MeasureWidthOrDefault(text, style, 0));
        double expectedColumn = Math.Truncate((glyphWidth * 1.22 + 5) / style.MaximumDigitWidthPixels * 256) / 256;
        Assert.InRange(Math.Abs(worksheet.Descendants<Row>().Single().Height!.Value - expectedHeight), 0, .02);
        Assert.InRange(Math.Abs(worksheet.GetFirstChild<Columns>()!.Elements<Column>().Single().Width!.Value - expectedColumn), 0, .02);
    }

    [Theory]
    [InlineData(45, 45d)]
    [InlineData(90, 90d)]
    [InlineData(160, -70d)]
    public void AutoFitPersistsProjectedMeasuredTextExtents(int rotation, double angle) {
        const string text = "Wide rotated text WWWWWW";
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Rotation");
        sheet.Cell(1, 1).SetValue(text).SetFontFamily("Calibri").SetTextRotation(rotation);
        sheet.AutoFitColumns();
        sheet.AutoFitRows();
        using var package = SpreadsheetDocument.Open(new MemoryStream(document.ToBytes()), false);
        var worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet;
        var measurer = OfficeTextMeasurer.Create(new OfficeFontInfo("Calibri", 11));
        var style = measurer.DefaultStyle;
        double width = measurer.MeasureWidthOrDefault(text, style, 0);
        double lineHeight = Math.Ceiling(measurer.MeasureLineHeightOrDefault("Xg", style, 0) + 2);
        double radians = Math.Abs(angle) * Math.PI / 180;
        double expectedPixels = width * Math.Cos(radians) + lineHeight * Math.Sin(radians);
        double expectedColumn = Math.Truncate((expectedPixels * 1.22 + 5) / style.MaximumDigitWidthPixels * 256) / 256;
        double unrotatedHeight = Math.Max(15d, lineHeight * .75);
        double expectedRow = Math.Max(15d, width * 1.22 * Math.Sin(radians) * .75 + unrotatedHeight * Math.Cos(radians));
        var column = worksheet.GetFirstChild<Columns>()!.Elements<Column>().Single();
        var row = worksheet.Descendants<Row>().Single();
        Assert.InRange(Math.Abs(column.Width!.Value - expectedColumn), 0, .02);
        Assert.InRange(Math.Abs(row.Height!.Value - expectedRow), 0, .02);
        var cell = Assert.Single(row.Elements<Cell>());
        var format = package.WorkbookPart.WorkbookStylesPart!.Stylesheet!.CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
        Assert.Equal((uint)rotation, format.Alignment!.TextRotation!.Value);
    }

    [Fact]
    public void AutoFitReservesThreeMeasuredSpacesPerIndentLevel() {
        using var source = new MemoryStream();
        using (var document = ExcelDocument.Create()) {
            var sheet = document.AddSheet("Indent");
            sheet.Cell(1, 1).SetValue("Indented text").SetFontFamily("Calibri");
            sheet.Cell(1, 2).SetValue("Indented text").SetFontFamily("Calibri");
            document.Save(source);
        }
        source.Position = 0;
        using (var package = SpreadsheetDocument.Open(source, true)) {
            var workbook = package.WorkbookPart!;
            var styles = workbook.WorkbookStylesPart!.Stylesheet!;
            var cell = workbook.WorksheetParts.Single().Worksheet.Descendants<Cell>().Single(item => item.CellReference!.Value == "B1");
            var format = (CellFormat)styles.CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value).CloneNode(true);
            format.Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Left, Indent = 4U };
            format.ApplyAlignment = true;
            cell.StyleIndex = (uint)styles.CellFormats.ChildElements.Count;
            styles.CellFormats.AppendChild(format);
            styles.CellFormats.Count = (uint)styles.CellFormats.ChildElements.Count;
        }
        source.Position = 0;
        using var loaded = ExcelDocument.Load(source);
        loaded.Sheets[0].AutoFitColumns();
        using var output = SpreadsheetDocument.Open(new MemoryStream(loaded.ToBytes()), false);
        var columns = output.WorkbookPart!.WorksheetParts.Single().Worksheet.GetFirstChild<Columns>()!.Elements<Column>().ToArray();
        double normal = columns.Single(column => column.Min!.Value == 1).Width!.Value;
        double indented = columns.Single(column => column.Min!.Value == 2).Width!.Value;
        var measurer = OfficeTextMeasurer.Create(new OfficeFontInfo("Calibri", 11));
        double expectedExtra = 12 * measurer.MeasureWidthOrDefault(" ", measurer.DefaultStyle, 0) * 1.22 / measurer.DefaultStyle.MaximumDigitWidthPixels;
        Assert.InRange(Math.Abs((indented - normal) - expectedExtra), 0, .02);
    }

    [Fact]
    public void AutoFitMeasuresFormattedTextSectionInsteadOfRawText() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Text");
        const string prefix = "A long displayed prefix: ";
        sheet.Cell(1, 1).SetValue("abc").SetNumberFormat("0;0;0;\"" + prefix + "\"@");
        sheet.Cell(1, 2).SetValue(prefix + "abc");
        sheet.AutoFitColumns();
        using var package = SpreadsheetDocument.Open(new MemoryStream(document.ToBytes()), false);
        var columns = package.WorkbookPart!.WorksheetParts.Single().Worksheet.GetFirstChild<Columns>()!.Elements<Column>().ToArray();
        double formatted = columns.Single(column => column.Min!.Value == 1).Width!.Value;
        double literal = columns.Single(column => column.Min!.Value == 2).Width!.Value;
        Assert.True(formatted > 15d);
        Assert.InRange(Math.Abs(formatted - literal), 0, .01);
    }
}
