using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Reader_XmlValueContentPreservesSplitTextAcrossReadPaths(bool useCachedFormulaResult, bool numericAsDecimal) {
        string path = CreateCompactFastPathWorkbook();
        try {
            using (var package = SpreadsheetDocument.Open(path, true)) {
                var part = package.WorkbookPart!.SharedStringTablePart
                    ?? package.WorkbookPart.AddNewPart<SharedStringTablePart>();
                part.SharedStringTable = new SharedStringTable(
                    Enumerable.Range(0, 12).Select(index => new SharedStringItem(new Text("String" + index))));
                part.SharedStringTable.Save();
            }
            string[] headers = { "Number", "Text", "Enabled", "Shared", "Long", "Next" };
            string headerXml = string.Concat(headers.Select((header, index) =>
                $"<c r=\"{(char)('A' + index)}1\" t=\"inlineStr\"><is><t>{header}</t></is></c>"));
            string xml = $$"""
                <?xml version="1.0" encoding="utf-16"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
                  <row r="1">{{headerXml}}</row>
                  <row r="2">
                    <c r="A2"><v>1<![CDATA[2]]>.5</v></c>
                    <c r="B2" t="str"><v>Al<![CDATA[pha]]>Beta</v></c>
                    <c r="C2" t="b"><v><![CDATA[1]]></v></c>
                    <c r="D2" t="s"><v>1<![CDATA[1]]></v></c>
                    <c r="E2"><v>{{new string('0', 80)}}1<![CDATA[2.5]]></v></c>
                    <c r="F2"><v>42</v></c>
                  </row>
                </sheetData></worksheet>
                """;
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml",
                Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(xml)).ToArray());
            var options = new ExcelReadOptions {
                UseCachedFormulaResult = useCachedFormulaResult,
                NumericAsDecimal = numericAsDecimal,
                InferDataTableColumnTypes = false
            };
            object number = numericAsDecimal ? (object)12.5m : 12.5d;
            object next = numericAsDecimal ? (object)42m : 42d;
            object[] expected = { number, "AlphaBeta", true, "String11", number, next };
            using (var reader = ExcelDocument.OpenDataReader(path, options)) {
                Assert.True(reader.Read());
                Assert.Equal(12.5d, reader.GetDouble(0));
                Assert.True(reader.GetBoolean(2));
                for (int column = 0; column < expected.Length; column++) Assert.Equal(expected[column], reader.GetValue(column));
                Assert.False(reader.Read());
            }
            using (var owner = ExcelDocumentReader.Open(path, options)) {
                var values = owner.GetSheet("Data").ReadRange("A2:F2", ExcelExecutionMode.Sequential);
                for (int column = 0; column < expected.Length; column++) Assert.Equal(expected[column], values[0, column]);
            }
            using (var owner = ExcelDocumentReader.Open(path, options))
            using (var table = owner.GetSheet("Data").ReadRangeAsDataTable("A1:F2", headersInFirstRow: true)) {
                Assert.Single(table.Rows.Cast<System.Data.DataRow>());
                for (int column = 0; column < expected.Length; column++) Assert.Equal(expected[column], table.Rows[0][column]);
            }
            using (var owner = ExcelDocumentReader.Open(path, options)) {
                var row = Assert.Single(owner.GetSheet("Data").ReadObjects<XmlValueTextRecord>("A1:F2"));
                Assert.Equal(12.5d, row.Number);
                Assert.Equal("AlphaBeta", row.Text);
                Assert.True(row.Enabled);
                Assert.Equal("String11", row.Shared);
                Assert.Equal(12.5d, row.Long);
                Assert.Equal(42, row.Next);
            }
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void Reader_EmptySharedStringValueAdvancesWithoutCachedResults() {
        string path = CreateCompactFastPathWorkbook();
        try {
            string xml = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
                + "<row r=\"1\"><c r=\"A1\" t=\"s\"><v/></c><c r=\"B1\"><v>42</v></c></row></sheetData></worksheet>";
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
            using var owner = ExcelDocumentReader.Open(path, new ExcelReadOptions { UseCachedFormulaResult = false });
            var values = owner.GetSheet("Data").ReadRange("A1:B1", ExcelExecutionMode.Sequential);
            Assert.Equal(string.Empty, values[0, 0]);
            Assert.Equal(42d, values[0, 1]);
        } finally {
            File.Delete(path);
        }
    }

    private sealed class XmlValueTextRecord {
        public double Number { get; set; }
        public string? Text { get; set; }
        public bool Enabled { get; set; }
        public string? Shared { get; set; }
        public double Long { get; set; }
        public int Next { get; set; }
    }
}
