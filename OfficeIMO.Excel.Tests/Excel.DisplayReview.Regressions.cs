using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public class ExcelDisplayReviewRegressionTests {
    [Fact]
    public void ImportedDisplayAndAutoFitUseCellThenRowThenColumnStylesIncludingExplicitZero() {
        using var source = new MemoryStream();
        using (var document = ExcelDocument.Create()) {
            foreach (string name in new[] { "Inherited", "Explicit" }) {
                var sheet = document.AddSheet(name);
                sheet.Cell(1, 1).SetValue(12.5).SetNumberFormat("0.00000000000000");
                sheet.Cell(1, 2).SetValue(46332d).SetNumberFormat("yyyy-mm-dd");
                sheet.Cell(1, 3).SetValue("abc").SetNumberFormat("0;0;0;\"A long prefix: \"@");
                sheet.Cell(1, 4).SetValue(12.5);
                sheet.Cell(2, 1).SetValue(12.5).SetNumberFormat("0.00").SetFontSize(global::OfficeIMO.OfficeLength.Points(24)).SetTextRotation(45);
                sheet.Cell(2, 2).SetValue(12.5).SetNumberFormat("0.00").SetFontSize(global::OfficeIMO.OfficeLength.Points(24)).SetTextRotation(45);
                sheet.Cell(2, 3).SetValue(12.5);
                sheet.Cell(3, 1).SetValue(12.5);
                sheet.Cell(4, 3).SetValue("abc\ndef").SetNumberFormat("0;0;0;\"A long prefix: \"@");
            }
            document.Save(source);
        }
        source.Position = 0;
        using (var package = SpreadsheetDocument.Open(source, true)) {
            var workbook = package.WorkbookPart!;
            string relationship = workbook.Workbook.Sheets!.Elements<Sheet>().Single(sheet => sheet.Name!.Value == "Inherited").Id!.Value!;
            var worksheet = ((WorksheetPart)workbook.GetPartById(relationship)).Worksheet;
            var cells = worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            uint columnStyle = cells["A1"].StyleIndex!.Value;
            uint dateStyle = cells["B1"].StyleIndex!.Value;
            uint textStyle = cells["C1"].StyleIndex!.Value;
            uint rowStyle = cells["A2"].StyleIndex!.Value;
            worksheet.InsertBefore(new Columns(
                new Column { Min = 1U, Max = 1U, Style = columnStyle },
                new Column { Min = 2U, Max = 2U, Style = dateStyle },
                new Column { Min = 3U, Max = 3U, Style = textStyle },
                new Column { Min = 4U, Max = 4U, Style = columnStyle }), worksheet.GetFirstChild<SheetData>()!);
            foreach (string address in new[] { "A1", "B1", "C1", "A2", "B2", "A3", "C4" }) cells[address].StyleIndex = null;
            cells["D1"].StyleIndex = 0U;
            cells["C2"].StyleIndex = 0U;
            var rows = worksheet.Descendants<Row>().ToDictionary(row => row.RowIndex!.Value);
            rows[2].StyleIndex = rowStyle;
            rows[2].CustomFormat = true;
            rows[3].StyleIndex = 0U;
            rows[3].CustomFormat = true;
        }
        source.Position = 0;
        using var loaded = ExcelDocument.Load(source);
        var inherited = loaded["Inherited"];
        AssertInheritedDisplay(inherited);
        foreach (var sheet in loaded.Sheets) {
            sheet.AutoFitColumns();
            sheet.AutoFitRows();
        }
        AssertInheritedDisplay(inherited);
        using var output = SpreadsheetDocument.Open(new MemoryStream(loaded.ToBytes()), false);
        var outputWorkbook = output.WorkbookPart!;
        Worksheet GetWorksheet(string name) => ((WorksheetPart)outputWorkbook.GetPartById(
            outputWorkbook.Workbook.Sheets!.Elements<Sheet>().Single(sheet => sheet.Name!.Value == name).Id!.Value!)).Worksheet!;
        var actual = GetWorksheet("Inherited");
        var expected = GetWorksheet("Explicit");
        for (uint column = 1; column <= 4; column++) {
            Assert.InRange(Math.Abs(ColumnWidth(actual, column) - ColumnWidth(expected, column)), 0, .01);
        }
        Assert.InRange(Math.Abs(actual.Descendants<Row>().Single(row => row.RowIndex!.Value == 2).Height!.Value
            - expected.Descendants<Row>().Single(row => row.RowIndex!.Value == 2).Height!.Value), 0, .02);
    }

    private static void AssertInheritedDisplay(ExcelSheet sheet) {
        Assert.Equal("12.50000000000000", sheet.Cell("A1").GetFormattedText(CultureInfo.InvariantCulture));
        Assert.Equal("2026-11-06", sheet.Cell("B1").GetFormattedText(CultureInfo.InvariantCulture));
        Assert.Equal("A long prefix: abc", sheet.Cell("C1").GetFormattedText(CultureInfo.InvariantCulture));
        Assert.Equal("A long prefix: abc\ndef", sheet.Cell("C4").GetFormattedText(CultureInfo.InvariantCulture));
        Assert.Equal("12.50", sheet.Cell("A2").GetFormattedText(CultureInfo.InvariantCulture));
        Assert.Equal("12.50", sheet.Cell("B2").GetFormattedText(CultureInfo.InvariantCulture));
        foreach (string address in new[] { "D1", "C2", "A3" })
            Assert.Equal("12.5", sheet.Cell(address).GetFormattedText(CultureInfo.InvariantCulture));
        Assert.Equal(12.5, sheet.Cell("A1").GetValue<double>());
        Assert.Equal(46332d, sheet.Cell("B1").GetValue<double>());
        Assert.Equal("abc", sheet.Cell("C1").GetValue<string>());
    }

    [Theory]
    [InlineData("0.00\\@", "12.50@")]
    [InlineData("0.00\"@\"", "12.50@")]
    [InlineData("\\@0.00", "@12.50")]
    [InlineData("\"@\"0.00", "@12.50")]
    [InlineData("0.00E+00\"@\"", "1.25E+01@")]
    [InlineData("\\@0.00E+00", "@1.25E+01")]
    [InlineData("\"@\"", "@")]
    public void NumericDisplayTreatsQuotedAndEscapedAtSignsAsLiterals(string format, string expected) {
        using var document = ExcelDocument.Create();
        document.AddSheet("Display").Cell(1, 1).SetValue(12.5).SetNumberFormat(format);
        using var reopened = ExcelDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal(expected, reopened.Sheets[0].Cell(1, 1).GetFormattedText(CultureInfo.InvariantCulture));
        Assert.Equal(12.5, reopened.Sheets[0].Cell(1, 1).GetValue<double>());
    }

    [Theory]
    [InlineData(ExcelDateSystem.NineteenHundred, "yyyy-mm-dd;yyyy-mm-dd;\"none\"", "2026-11-06")]
    [InlineData(ExcelDateSystem.NineteenFour, "yyyy-mm-dd;yyyy-mm-dd;\"none\"", "2026-11-06")]
    [InlineData(ExcelDateSystem.NineteenHundred, "[>=45000]yyyy-mm-dd;\"none\"", "2026-11-06")]
    [InlineData(ExcelDateSystem.NineteenFour, "[>=45000]yyyy-mm-dd;\"none\"", "none")]
    [InlineData(ExcelDateSystem.NineteenFour, "[>=44000]yyyy-mm-dd;\"none\"", "2026-11-06")]
    public void IsoDateSelectsSectionsUsingTheActualWorkbookSerial(ExcelDateSystem system, string format, string expected) {
        using var source = CreateIsoDateWorkbook("2026-11-06T13:00:00", format, system);
        using var loaded = ExcelDocument.Load(source);
        Assert.Equal(expected, loaded.Sheets[0].Cell(1, 1).GetFormattedText(CultureInfo.InvariantCulture));
        Assert.Equal("2026-11-06T13:00:00", loaded.Sheets[0].Cell(1, 1).GetValue<string>());
    }

    [Fact]
    public void IsoBuiltInShortDateUsesTheRequestedCulture() {
        using var source = CreateIsoDateWorkbook("2026-11-06T13:00:00", "m/d/yyyy", ExcelDateSystem.NineteenFour, builtInShortDate: true);
        using var loaded = ExcelDocument.Load(source);
        Assert.Equal("2026/11/06", loaded.Sheets[0].Cell(1, 1).GetFormattedText(CultureInfo.GetCultureInfo("ja-JP")));
    }

    [Fact]
    public void LocalizedDateAutoFitUsesACalendarRepresentableSample() {
        var date = new DateTime(2026, 11, 6);
        var culture = CultureInfo.GetCultureInfo("ar-SA");
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Dates");
        sheet.Cell(1, 1).SetValue(date).SetNumberFormat("[$-401]d-mmm");
        Assert.Equal(date.ToString("%d", culture) + "-" + date.ToString("MMM", culture),
            sheet.Cell(1, 1).GetFormattedText(CultureInfo.InvariantCulture));
        sheet.AutoFitColumns();
        using var output = SpreadsheetDocument.Open(new MemoryStream(document.ToBytes()), false);
        Assert.InRange(ColumnWidth(output.WorkbookPart!.WorksheetParts.Single().Worksheet, 1), 1, 255);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void DatesOutsideTheProviderCalendarRetainFallbackText(bool iso, bool builtIn) {
        const string raw = "2099-12-31T13:00:00";
        var provider = CultureInfo.GetCultureInfo("ar-SA");
        string format = builtIn ? "m/d/yyyy" : "[$-401]yyyy-mm-dd";
        using var source = CreateIsoDateWorkbook(raw, format, ExcelDateSystem.NineteenHundred, builtIn);
        double serial = ExcelDateSystemConverter.ToSerial(new DateTime(2099, 12, 31, 13, 0, 0), ExcelDateSystem.NineteenHundred);
        if (!iso) {
            using (var package = SpreadsheetDocument.Open(source, true)) {
                var cell = package.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>().Single();
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(serial.ToString("R", CultureInfo.InvariantCulture));
            }
            source.Position = 0;
        }
        using var loaded = ExcelDocument.Load(source);
        Assert.Equal(iso ? raw : serial.ToString(provider), loaded.Sheets[0].Cell(1, 1).GetFormattedText(provider));
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, true, false)]
    [InlineData(false, false, true)]
    [InlineData(true, false, true)]
    public void TextFormatKeepsRichRunMetricsForInlineAndSharedStrings(bool shared, bool repeat, bool suppress) {
        const string raw = "WWWWiiii";
        var valueRuns = new[] {
            new ExcelRichTextRun("WWWW") { FontSize = global::OfficeIMO.OfficeLength.Points(40), FontFamily = "Calibri" },
            new ExcelRichTextRun("iiii") { FontSize = global::OfficeIMO.OfficeLength.Points(8), FontFamily = "Arial" }
        };
        string format = suppress ? "0;0;0;\"shown\"" : repeat ? "0;0;0;\"x\"@\"/\"@\"y\"" : "0;0;0;\"x\"@\"y\"";
        var displayed = new List<ExcelRichTextRun>();
        if (suppress) displayed.Add(new ExcelRichTextRun("shown"));
        else {
            displayed.Add(new ExcelRichTextRun("x"));
            displayed.AddRange(valueRuns);
            if (repeat) { displayed.Add(new ExcelRichTextRun("/")); displayed.AddRange(valueRuns); }
            displayed.Add(new ExcelRichTextRun("y"));
        }
        using var source = new MemoryStream();
        using (var document = ExcelDocument.Create()) {
            var sheet = document.AddSheet("Rich");
            sheet.Cell(1, 1).SetRichText(valueRuns).SetNumberFormat(format);
            sheet.Cell(1, 2).SetRichText(displayed.ToArray());
            document.Save(source);
        }
        source.Position = 0;
        if (shared) {
            using (var package = SpreadsheetDocument.Open(source, true)) {
                var workbook = package.WorkbookPart!;
                var part = workbook.SharedStringTablePart ?? workbook.AddNewPart<SharedStringTablePart>();
                part.SharedStringTable ??= new SharedStringTable();
                foreach (var cell in workbook.WorksheetParts.Single().Worksheet.Descendants<Cell>()) {
                    uint index = (uint)part.SharedStringTable.ChildElements.Count;
                    part.SharedStringTable.AppendChild(new SharedStringItem(cell.InlineString!.ChildElements.Select(child => child.CloneNode(true))));
                    cell.InlineString = null;
                    cell.DataType = CellValues.SharedString;
                    cell.CellValue = new CellValue(index.ToString(CultureInfo.InvariantCulture));
                }
                part.SharedStringTable.Count = (uint)part.SharedStringTable.ChildElements.Count;
                part.SharedStringTable.UniqueCount = (uint)part.SharedStringTable.ChildElements.Count;
            }
            source.Position = 0;
        }
        using var loaded = ExcelDocument.Load(source);
        var loadedSheet = loaded.Sheets[0];
        loadedSheet.AutoFitColumns();
        Assert.Equal(suppress ? "shown" : repeat ? "x" + raw + "/" + raw + "y" : "x" + raw + "y",
            loadedSheet.Cell(1, 1).GetFormattedText(CultureInfo.InvariantCulture));
        Assert.Equal(raw, loadedSheet.Cell(1, 1).GetValue<string>());
        using var output = SpreadsheetDocument.Open(new MemoryStream(loaded.ToBytes()), false);
        var worksheet = output.WorkbookPart!.WorksheetParts.Single().Worksheet;
        Assert.InRange(Math.Abs(ColumnWidth(worksheet, 1) - ColumnWidth(worksheet, 2)), 0, .01);
    }

    private static MemoryStream CreateIsoDateWorkbook(string raw, string format, ExcelDateSystem system, bool builtInShortDate = false) {
        var source = new MemoryStream();
        using (var document = ExcelDocument.Create()) {
            document.DateSystem = system;
            document.AddSheet("Dates").Cell(1, 1).SetValue(0d).SetNumberFormat(format);
            document.Save(source);
        }
        source.Position = 0;
        using (var package = SpreadsheetDocument.Open(source, true)) {
            var workbook = package.WorkbookPart!;
            var cell = workbook.WorksheetParts.Single().Worksheet.Descendants<Cell>().Single();
            cell.DataType = CellValues.Date;
            cell.CellValue = new CellValue(raw);
            if (builtInShortDate) {
                var styles = workbook.WorkbookStylesPart!.Stylesheet!;
                styles.CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value).NumberFormatId = 14U;
            }
        }
        source.Position = 0;
        return source;
    }

    private static double ColumnWidth(Worksheet worksheet, uint index) => worksheet.GetFirstChild<Columns>()!
        .Elements<Column>().Single(column => column.Min!.Value <= index && column.Max!.Value >= index).Width!.Value;

    [Theory]
    [InlineData("h:mm", false, "12:00")]
    [InlineData("[$-401]h:mm", true, "12:00")]
    [InlineData("h:mm AM/PM", false, "12:00 PM")]
    [InlineData("h:mm\" d\"", false, "12:00 d")]
    public void TimeOnlyDisplayDoesNotRequireTheAnchorDateToFitTheLocaleCalendar(string format, bool invariant, string expected) {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Clock");
        sheet.Cell(1, 1).SetValue(0.5).SetNumberFormat(format);
        var provider = invariant ? CultureInfo.InvariantCulture : CultureInfo.GetCultureInfo("ar-SA");
        Assert.Equal(expected, sheet.Cell(1, 1).GetFormattedText(provider));
        sheet.AutoFitColumns();
        Assert.Equal(expected, sheet.Cell(1, 1).GetFormattedText(provider));
        Assert.Equal(0.5, sheet.Cell(1, 1).GetValue<double>());
    }

}
