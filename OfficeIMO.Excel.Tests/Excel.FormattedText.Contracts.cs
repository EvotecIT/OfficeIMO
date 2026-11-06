using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public class ExcelFormattedTextContractTests {
    [Theory]
    [InlineData(0.30000000000000004, "0,3")]
    [InlineData(0.00001, "0,00001")]
    [InlineData(1E21, "1E+21")]
    public void GeneralFormattedTextUsesCanonicalPrecisionAndCurrentCulture(double value, string expected) {
        var original = CultureInfo.CurrentCulture;
        try {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("nl-NL");
            using var document = ExcelDocument.Create();
            var cell = document.AddSheet("General").Cell(1, 1).SetValue(value);
            Assert.Equal(expected, cell.GetFormattedText());
            Assert.Equal(value, cell.GetValue<double>());
        } finally {
            CultureInfo.CurrentCulture = original;
        }
    }

    [Theory]
    [InlineData(2.5, "0", "en-US", "3")]
    [InlineData(-2.5, "0", "en-US", "-3")]
    [InlineData(0.125, "0.00", "en-US", "0.13")]
    [InlineData(1234.56, "#,##0.00", "nl-NL", "1.234,56")]
    [InlineData(12.5, "0.0000", "nl-NL", "12,5000")]
    [InlineData(1.2, "000.0#", "nl-NL", "001,2")]
    [InlineData(1234.56, "\",. \"#,##0.00\" .,;\"", "nl-NL", ",. 1.234,56 .,;")]
    [InlineData(46332, "dd/mm/yyyy", "en-US", "06/11/2026")]
    [InlineData(46332, "yyyy-mm-dd", "pl-PL", "2026-11-06")]
    [InlineData(46332, "[$-409]d-mmm", "pl-PL", "6-Nov")]
    [InlineData(46332, "d-mmmm", "nl-NL", "6-november")]
    public void FormattedTextAppliesPersistedNumberFormatWithoutChangingRawValue(double value, string format, string culture, string expected) {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Display");
        sheet.Cell(1, 1).SetValue(value).SetNumberFormat(format);
        Assert.Equal(expected, sheet.Cell(1, 1).GetFormattedText(CultureInfo.GetCultureInfo(culture)));
        Assert.Equal(value, sheet.Cell(1, 1).GetValue<double>());
        byte[] bytes = document.ToBytes();
        using (var package = SpreadsheetDocument.Open(new MemoryStream(bytes), false)) {
            var cell = package.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>().Single();
            Assert.Equal(value, double.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture));
            var styles = package.WorkbookPart.WorkbookStylesPart!.Stylesheet!;
            uint id = styles.CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value).NumberFormatId!.Value;
            Assert.Equal(format, styles.NumberingFormats!.Elements<NumberingFormat>().Single(item => item.NumberFormatId!.Value == id).FormatCode!.Value);
        }
        using var reopened = ExcelDocument.Load(new MemoryStream(bytes));
        Assert.Equal(expected, reopened.Sheets[0].Cell(1, 1).GetFormattedText(CultureInfo.GetCultureInfo(culture)));
    }

    [Theory]
    [InlineData("0;0;0;\"Prefix: \"@", "Prefix: abc")]
    [InlineData("\"<\"@\">\"", "<abc>")]
    [InlineData("0;0;0;\"constant\"", "constant")]
    [InlineData("0;0;0;", "")]
    [InlineData("0;0;0;[Unknown]@", "[Unknown]abc")]
    public void FormattedTextAppliesTextSectionAndPreservesUnknownTokens(string format, string expected) {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Display");
        sheet.Cell(1, 1).SetValue("abc").SetNumberFormat(format);
        using var reopened = ExcelDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal(expected, reopened.Sheets[0].Cell(1, 1).GetFormattedText(CultureInfo.InvariantCulture));
        Assert.Equal("abc", reopened.Sheets[0].Cell(1, 1).GetValue<string>());
    }

    [Fact]
    public void FormattedTextUsesNumericFormulaCacheWithoutRecalculation() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Display");
        sheet.Cell(1, 1).SetValue(1d);
        sheet.Cell(1, 2).SetValue(1d).SetNumberFormat("0.00");
        using var source = new MemoryStream();
        document.Save(source);
        source.Position = 0;
        using (var package = SpreadsheetDocument.Open(source, true)) {
            var cell = package.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>().Single(item => item.CellReference!.Value == "B1");
            cell.CellFormula = new CellFormula("SUM(A1)");
            cell.CellValue = new CellValue("1");
            cell.DataType = CellValues.Number;
        }
        source.Position = 0;
        using var reopened = ExcelDocument.Load(source);
        var formula = reopened.Sheets[0].Cell(1, 2);
        Assert.Equal("1.00", formula.GetFormattedText(CultureInfo.InvariantCulture));
        Assert.Equal(1d, Assert.IsType<double>(formula.GetValue().Value));
        Assert.Equal("1", formula.GetValue().CachedText);
    }

    [Theory]
    [InlineData("dd/mm/yyyy", "06/11/2026")]
    [InlineData("General", "2026-11-06T13:00:00")]
    public void FormattedTextAppliesDateFormatToImportedIsoDateCell(string format, string expected) {
        using var document = ExcelDocument.Create();
        document.AddSheet("Dates").Cell(1, 1).SetValue(0d).SetNumberFormat(format);
        using var source = new MemoryStream();
        document.Save(source);
        source.Position = 0;
        using (var package = SpreadsheetDocument.Open(source, true)) {
            var cell = package.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>().Single();
            cell.DataType = CellValues.Date;
            cell.CellValue = new CellValue("2026-11-06T13:00:00");
        }
        source.Position = 0;
        using var reopened = ExcelDocument.Load(source);
        Assert.Equal(expected, reopened.Sheets[0].Cell(1, 1).GetFormattedText(CultureInfo.GetCultureInfo("en-US")));
        Assert.Equal("2026-11-06T13:00:00", reopened.Sheets[0].Cell(1, 1).GetValue<string>());
    }

    [Fact]
    public void BuiltInShortDateUsesRequestedCultureAndWorkbookDateSystem() {
        using var document = ExcelDocument.Create();
        document.DateSystem = ExcelDateSystem.NineteenFour;
        var sheet = document.AddSheet("Dates");
        sheet.Cell(1, 1).SetValue(44870d).SetNumberFormat("m/d/yyyy");
        using var source = new MemoryStream();
        document.Save(source);
        source.Position = 0;
        using (var package = SpreadsheetDocument.Open(source, true)) {
            var cell = package.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>().Single();
            var styles = package.WorkbookPart.WorkbookStylesPart!.Stylesheet!;
            var format = styles.CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
            format.NumberFormatId = 14U;
        }
        source.Position = 0;
        using var reopened = ExcelDocument.Load(source);
        Assert.Equal("2026/11/06", reopened.Sheets[0].Cell(1, 1).GetFormattedText(CultureInfo.GetCultureInfo("ja-JP")));
        Assert.Equal(44870d, reopened.Sheets[0].Cell(1, 1).GetValue<double>());
    }

}
