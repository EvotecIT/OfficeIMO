using System.Data;
using System.Text;
using System.Threading;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    public static IEnumerable<object[]> CompletedRangeReplayCases() {
        foreach (string api in new[] { "RangeSequential", "RangeParallel", "Rows", "Column", "StreamSingle", "StreamChunks", "Dictionaries", "DictionariesSequential", "Objects", "ObjectsSequential", "DataTable", "InferredDataTable", "Cells", "DataReader", "DataReaderDom", "TypedObjectsStream" })
            foreach (bool utf16 in new[] { false, true })
                foreach (bool cancellable in new[] { false, true })
                    yield return new object[] { api, utf16, cancellable };
    }

    [Theory]
    [MemberData(nameof(CompletedRangeReplayCases))]
    public void Reader_CompletedRangeStillReadsLaterPhysicalRows(string api, bool utf16, bool cancellable) {
        string path = CreateCompactFastPathWorkbook();
        string encodingName = utf16 ? "utf-16" : "utf-8";
        string outsideRows = string.Concat(Enumerable.Range(4, 17).Select(row => $"<row r=\"{row}\"><c r=\"A{row}\" t=\"str\"><v>outside</v></c></row>")) + "<row r=\"21\"/>";
        string xml = $$"""
            <?xml version="1.0" encoding="{{encodingName}}"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
              <row r="1"><c r="A1" t="str"><v>Value</v></c></row>
              <row r="2"><c r="A2" t="str"><v>old</v></c></row>
              <row r="3"><c r="A3" t="str"><v>second</v></c></row>
              {{outsideRows}}
              <row r="2"><c r="A2" t="str"><v>new</v></c></row>
            </sheetData></worksheet>
            """;
        try {
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", (utf16 ? Encoding.Unicode : Encoding.UTF8).GetBytes(xml));
            using var cancellation = new CancellationTokenSource();
            CancellationToken token = cancellable ? cancellation.Token : default;
            var options = new ExcelReadOptions { InferDataTableColumnTypes = api == "InferredDataTable" };
            if (api == "DataReaderDom") options.Culture = System.Globalization.CultureInfo.GetCultureInfo("fr-FR");
            using var owner = ExcelDocumentReader.Open(path, options);
            var sheet = owner.GetSheet("Data");
            object?[] actual;
            switch (api) {
                case "RangeSequential":
                case "RangeParallel": {
                    var values = sheet.ReadRange("A2:A3", api == "RangeParallel" ? ExcelExecutionMode.Parallel : ExcelExecutionMode.Sequential, token);
                    actual = new[] { values[0, 0], values[1, 0] };
                    break;
                }
                case "Rows": actual = sheet.ReadRows("A2:A3", token).Select(row => row?[0]).ToArray(); break;
                case "Column": actual = sheet.ReadColumn("A2:A3", token).ToArray(); break;
                case "DataReader":
                case "DataReaderDom": {
                    using var reader = sheet.ReadRangeAsDataReader("A1:A3", chunkRows: 1, schemaSampleRows: 0, ct: token);
                    var values = new List<object?>();
                    while (reader.Read()) values.Add(reader.GetValue(0));
                    actual = values.ToArray();
                    break;
                }
                case "StreamSingle":
                case "StreamChunks": {
                    actual = new object?[2];
                    foreach (var chunk in sheet.ReadRangeStream("A2:A3", chunkRows: api == "StreamSingle" ? 2 : 1, ct: token))
                        for (int row = 0; row < chunk.RowCount; row++) actual[chunk.StartRow + row - 2] = chunk.Rows[row][0];
                    break;
                }
                case "Dictionaries":
                case "DictionariesSequential":
                    actual = sheet.ReadObjects("A1:A3", mode: api == "DictionariesSequential" ? ExcelExecutionMode.Sequential : ExcelExecutionMode.Automatic, ct: token).Select(row => row["Value"]).ToArray();
                    break;
                case "Objects":
                case "ObjectsSequential":
                    actual = sheet.ReadObjects<CompletedRangeReplayRow>("A1:A3", api == "ObjectsSequential" ? ExcelExecutionMode.Sequential : ExcelExecutionMode.Automatic, token).Select(row => (object?)row.Value).ToArray();
                    break;
                case "TypedObjectsStream":
                    actual = sheet.ReadObjectsStream<CompletedRangeReplayRow>("A1:A3", token).Select(row => (object?)row.Value).ToArray();
                    break;
                case "DataTable":
                case "InferredDataTable": {
                    using DataTable table = sheet.ReadRangeAsDataTable("A1:A3", headersInFirstRow: true, ct: token);
                    actual = table.Rows.Cast<DataRow>().Select(row => row[0]).ToArray();
                    break;
                }
                case "Cells": {
                    actual = new object?[2];
                    var cells = sheet.EnumerateRange("A2:A3", token).ToArray();
                    Assert.Equal(3, cells.Length);
                    foreach (var cell in cells) actual[cell.Row - 2] = cell.Value;
                    break;
                }
                default: throw new InvalidOperationException(api);
            }
            Assert.Equal(new object?[] { "new", "second" }, actual);
        } finally {
            File.Delete(path);
        }
    }

    public sealed class CompletedRangeReplayRow {
        public string? Value { get; set; }
    }
}
