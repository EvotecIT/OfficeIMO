using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public class ExcelTextEscapingTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void TabularExportPreservesTextAndSanitizesControls(bool sharedStrings, bool cellReferences) {
        const string header = "Notes & <tag> '\"";
        string[] values = Values().ToArray();
        using var table = new DataTable();
        table.Columns.Add(header, typeof(string));
        foreach (string value in values) table.Rows.Add(value);
        using var stream = new MemoryStream();
        using (var reader = table.CreateDataReader()) {
            ExcelDocument.WriteDataReader(stream, reader, new ExcelTabularWriteOptions {
                SheetName = "Data & '\"",
                UseSharedStrings = sharedStrings,
                IncludeCellReferences = cellReferences
            });
        }
        stream.Position = 0;
        using var package = SpreadsheetDocument.Open(stream, false);
        WorkbookPart workbook = package.WorkbookPart!;
        string[] strings = workbook.SharedStringTablePart?.SharedStringTable?
            .Elements<SharedStringItem>().Select(item => item.InnerText).ToArray() ?? Array.Empty<string>();
        string[] actual = workbook.WorksheetParts.Single().Worksheet.Descendants<Cell>().Select(cell =>
            cell.DataType?.Value == CellValues.SharedString
                ? strings[int.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture)]
                : cell.InlineString?.InnerText ?? cell.CellValue?.Text ?? "").ToArray();
        string[] expected = new[] { header }.Concat(values.Select(value =>
            new string(value.Where(c => c >= 0x20 || c == '\t' || c == '\n' || c == '\r').ToArray()))).ToArray();
        Assert.Equal(expected, actual);
        Assert.Equal("Data & '\"", workbook.Workbook.Sheets!.Elements<Sheet>().Single().Name!.Value);
        if (cellReferences) {
            stream.Position = 0;
            using var reopened = ExcelDocument.Load(stream);
            for (int row = 1; row <= expected.Length; row++) {
                Assert.True(reopened["Data & '\""].TryGetCellText(row, 1, out string text));
                Assert.Equal(expected[row - 1], text);
            }
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CellValuesPreserveCarriageReturnsAcrossPackageWriters(bool standardWriter) {
        const string value = "Start\r\nMiddle\rEnd\nTab\tUnicode Łódź 😀 &<>";
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(new MemoryStream())) {
            var sheet = document.AddWorksheet("Text");
            sheet.CellValue(1, 1, value);
            sheet.CellValue(1, 2, "inline");
            sheet.CellValue(1, 3, "plain");
            var cells = sheet.WorksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            cells["B1"].CellValue = null;
            cells["B1"].DataType = CellValues.InlineString;
            cells["B1"].InlineString = new InlineString(new Text(value));
            cells["C1"].DataType = CellValues.String;
            cells["C1"].CellValue = new CellValue(value);
            sheet.MarkRequiresSavePreparation();
            document.Save(stream, new ExcelSaveOptions { DisableFastPackageWriter = standardWriter });
            Assert.Equal(standardWriter ? ExcelSavePackageWriter.StandardPackage : ExcelSavePackageWriter.SimplePackage,
                document.LastSaveDiagnostics.Writer);
        }
        stream.Position = 0;
        using (var package = SpreadsheetDocument.Open(stream, false)) {
            var writtenCells = package.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>()
                .ToDictionary(cell => cell.CellReference!.Value!);
            var cell = writtenCells["A1"];
            int index = int.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture);
            Assert.Equal(value, package.WorkbookPart.SharedStringTablePart!.SharedStringTable!
                .Elements<SharedStringItem>().ElementAt(index).InnerText);
            Assert.Equal(value, writtenCells["B1"].InlineString!.InnerText);
            Assert.Equal(value, writtenCells["C1"].CellValue!.Text);
        }

        stream.Position = 0;
        using var reopened = ExcelDocument.Load(stream);
        for (int column = 1; column <= 3; column++) {
            Assert.True(reopened["Text"].TryGetCellText(1, column, out string actual));
            Assert.Equal(value, actual);
        }
    }

    [Fact]
    public void ExtendedPackagePreservesCarriageReturnsInSharedStrings() {
        const string value = "First\r\nSecond\rThird";
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(new MemoryStream())) {
            var sheet = document.AddWorksheet("Text");
            sheet.CellValue(1, 1, value);
            sheet.AddChart(
                new ExcelChartData(new[] { "One" }, new[] { new ExcelChartSeries("Values", new[] { 1.0 }) }),
                row: 2, column: 3, type: ExcelChartType.ColumnClustered, title: "Values");
            document.Save(stream);
            Assert.Equal(ExcelSavePackageWriter.ExtendedPackage, document.LastSaveDiagnostics.Writer);
        }

        stream.Position = 0;
        using var reopened = ExcelDocument.Load(stream);
        Assert.True(reopened["Text"].TryGetCellText(1, 1, out string actual));
        Assert.Equal(value, actual);
    }

    [Fact]
    public void ExtendedPackageFallbackPreservesCarriageReturnsInWorksheetCells() {
        const string value = "First\r\nSecond\rThird\nFourth";
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(new MemoryStream())) {
            var sheet = document.AddWorksheet("Text");
            sheet.CellValue(1, 1, value);
            sheet.CellValue(1, 2, "inline");
            sheet.CellValue(1, 3, "plain");
            sheet.CellValue(1, 4, "#DIV/0!");
            var cells = sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                .ToDictionary(cell => cell.CellReference!.Value!);
            cells["B1"].CellValue = null;
            cells["B1"].DataType = CellValues.InlineString;
            cells["B1"].InlineString = new InlineString(new Text(value));
            cells["C1"].DataType = CellValues.String;
            cells["C1"].CellValue = new CellValue(value);
            cells["D1"].DataType = CellValues.Error;
            sheet.AddChart(
                new ExcelChartData(new[] { "One" }, new[] { new ExcelChartSeries("Values", new[] { 1.0 }) }),
                row: 2, column: 3, type: ExcelChartType.ColumnClustered, title: "Values");
            sheet.MarkRequiresSavePreparation();
            document.Save(stream);
            Assert.Equal(ExcelSavePackageWriter.ExtendedPackage, document.LastSaveDiagnostics.Writer);
        }

        stream.Position = 0;
        using var package = SpreadsheetDocument.Open(stream, false);
        var textSheetId = package.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>()
            .Single(sheet => sheet.Name!.Value == "Text").Id!.Value!;
        var worksheetPart = (WorksheetPart)package.WorkbookPart.GetPartById(textSheetId);
        var writtenCells = worksheetPart.Worksheet.Descendants<Cell>()
            .ToDictionary(cell => cell.CellReference!.Value!);
        Assert.Equal(value, writtenCells["B1"].InlineString!.InnerText);
        Assert.Equal(value, writtenCells["C1"].CellValue!.Text);
        Assert.Equal(CellValues.Error, writtenCells["D1"].DataType!.Value);

        stream.Position = 0;
        using var reopened = ExcelDocument.Load(stream);
        foreach (int column in new[] { 1, 2, 3 }) {
            Assert.True(reopened["Text"].TryGetCellText(1, column, out string actual));
            Assert.Equal(value, actual);
        }
    }

    [Fact]
    public void StandardPackagePreservesCarriageReturnsAcrossMultipleWorksheets() {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(new MemoryStream())) {
            for (int index = 0; index < 12; index++) {
                var sheet = document.AddWorksheet("Text" + index);
                sheet.CellValue(1, 1, "Sheet" + index + "\r\nLine\rEnd");
                var cell = sheet.WorksheetPart.Worksheet.Descendants<Cell>().Single();
                cell.CellValue = null;
                cell.DataType = CellValues.InlineString;
                cell.InlineString = new InlineString(new Text("Sheet" + index + "\r\nLine\rEnd"));
                sheet.MarkRequiresSavePreparation();
            }
            document.Save(stream, new ExcelSaveOptions { DisableFastPackageWriter = true });
            Assert.Equal(ExcelSavePackageWriter.StandardPackage, document.LastSaveDiagnostics.Writer);
        }

        stream.Position = 0;
        using (var package = SpreadsheetDocument.Open(stream, false)) {
            foreach (var sheet in package.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>()) {
                int index = int.Parse(sheet.Name!.Value!.Substring(4), CultureInfo.InvariantCulture);
                var part = (WorksheetPart)package.WorkbookPart.GetPartById(sheet.Id!.Value!);
                var cell = part.Worksheet.Descendants<Cell>().Single();
                string actual = cell.InlineString!.InnerText;
                Assert.Equal("Sheet" + index + "\r\nLine\rEnd", actual);
            }
        }

        stream.Position = 0;
        using var reopened = ExcelDocument.Load(stream);
        for (int index = 0; index < 12; index++) {
            Assert.True(reopened["Text" + index].TryGetCellText(1, 1, out string actual));
            Assert.Equal("Sheet" + index + "\r\nLine\rEnd", actual);
        }
    }

    private static IEnumerable<string> Values() {
        yield return "";
        yield return " \tŁódź 😀 &<> '\"\n ";
        yield return "before\r\nafter\rmiddle\nlast";
        yield return new string(Enumerable.Range(0, 32).Select(value => (char)value).ToArray()) + "valid";
        yield return string.Concat(Enumerable.Repeat("&<>\u0001", 1000));
        foreach (int offset in new[] { 0, 15, 16, 31, 32, 63, 64, 4096, 32000 }) {
            yield return new string('a', offset) + "&<\u0001>tail";
            yield return "<&>" + new string('b', offset);
        }
    }
}
