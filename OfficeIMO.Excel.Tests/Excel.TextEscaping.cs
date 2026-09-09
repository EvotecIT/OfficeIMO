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
            new string(value.Where(c => c >= 0x20 || c == '\t' || c == '\n' || c == '\r').ToArray())
                .Replace("\r\n", "\n").Replace('\r', '\n'))).ToArray();
        Assert.Equal(expected, actual);
        Assert.Equal("Data & '\"", workbook.Workbook.Sheets!.Elements<Sheet>().Single().Name!.Value);
    }

    private static IEnumerable<string> Values() {
        yield return "";
        yield return " \tŁódź 😀 &<> '\"\n ";
        yield return new string(Enumerable.Range(0, 32).Select(value => (char)value).ToArray()) + "valid";
        yield return string.Concat(Enumerable.Repeat("&<>\u0001", 1000));
        foreach (int offset in new[] { 0, 15, 16, 31, 32, 63, 64, 4096, 32000 }) {
            yield return new string('a', offset) + "&<\u0001>tail";
            yield return "<&>" + new string('b', offset);
        }
    }
}
