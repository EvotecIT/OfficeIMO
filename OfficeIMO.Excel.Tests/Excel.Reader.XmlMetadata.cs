using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData("plain", false)]
    [InlineData("plain", true)]
    [InlineData("entities", false)]
    [InlineData("entities", true)]
    [InlineData("padded-style", false)]
    [InlineData("padded-style", true)]
    [InlineData("long-style", false)]
    [InlineData("long-style", true)]
    [InlineData("unknown-type", false)]
    [InlineData("unknown-type", true)]
    public void Reader_XmlMetadataPreservesTypedAndMaterializedValues(string shape, bool date1904) {
        string path = CreateCompactFastPathWorkbook();
        try {
            using (var workbook = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(path, true)) {
                workbook.WorkbookPart!.Workbook.WorkbookProperties =
                    new DocumentFormat.OpenXml.Spreadsheet.WorkbookProperties { Date1904 = date1904 };
            }
            string numberType = shape == "entities" ? "&#110;" : shape == "unknown-type" ? new string('x', 31) + "😀tail" : "n";
            string booleanType = shape == "entities" ? "&#98;" : "b";
            string style = shape switch {
                "entities" => "&#49;",
                "padded-style" => " +1 ",
                "long-style" => new string('0', 80) + "1",
                _ => "1"
            };
            string xml = $$"""
                <x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData>
                  <x:row r="1"><x:c r="A1" t="inlineStr"><x:is><x:t>Id</x:t></x:is></x:c><x:c r="B1" t="inlineStr"><x:is><x:t>Active</x:t></x:is></x:c><x:c r="C1" t="inlineStr"><x:is><x:t>Name</x:t></x:is></x:c><x:c r="D1" t="inlineStr"><x:is><x:t>Created</x:t></x:is></x:c></x:row>
                  <x:row r="2"><x:c r="A2" t="{{numberType}}" s="0"><x:v>42</x:v></x:c><x:c r="B2" t="{{booleanType}}"><x:v>1</x:v></x:c><x:c r="C2" t="str"><x:v>Alpha</x:v></x:c><x:c r="D2" t="n" s="{{style}}"><x:v>45351</x:v></x:c></x:row>
                  <x:row r="3"><x:c r="A3" t="n" s="0"><x:v>43</x:v></x:c><x:c r="B3" t="b"><x:v>0</x:v></x:c><x:c r="C3" t="str"><x:v>Beta</x:v></x:c><x:c r="D3" s="1"><x:v>45352</x:v></x:c></x:row>
                </x:sheetData></x:worksheet>
                """;
            ReplaceXmlMetadataWorksheet(path, xml);
            ReplaceZipEntry(path, "xl/styles.xml", Encoding.UTF8.GetBytes("""
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs></styleSheet>
                """));
            var options = new ExcelReadOptions { NumericAsDecimal = true, InferDataTableColumnTypes = false };
            object[][] expected = [
                [42m, true, "Alpha", new DateTime(2024, 2, 29).AddDays(date1904 ? 1462 : 0)],
                [43m, false, "Beta", new DateTime(2024, 3, 1).AddDays(date1904 ? 1462 : 0)]
            ];
            using (var reader = ExcelDocument.OpenDataReader(path, options)) {
                for (int row = 0; row < 2; row++) {
                    Assert.True(reader.Read());
                    Assert.Equal(expected[row][0], reader.GetDecimal(0));
                    Assert.Equal(expected[row][1], reader.GetBoolean(1));
                    Assert.Equal(expected[row][2], reader.GetString(2));
                    Assert.Equal(expected[row][3], reader.GetDateTime(3));
                    for (int column = 0; column < 4; column++) Assert.Equal(expected[row][column], reader.GetValue(column));
                }
                Assert.False(reader.Read());
            }
            using var owner = ExcelDocumentReader.Open(path, options);
            var sheet = owner.GetSheet("Data");
            var values = sheet.ReadRange("A2:D3", ExcelExecutionMode.Sequential);
            using var table = sheet.ReadRangeAsDataTable("A1:D3", headersInFirstRow: true);
            var objects = sheet.ReadObjects<XmlMetadataRecord>("A1:D3").ToArray();
            Assert.Equal(2, table.Rows.Count);
            Assert.Equal(2, objects.Length);
            for (int row = 0; row < 2; row++) {
                for (int column = 0; column < 4; column++) {
                    Assert.Equal(expected[row][column], values[row, column]);
                    Assert.Equal(expected[row][column], table.Rows[row][column]);
                }
                Assert.Equal(expected[row][0], objects[row].Id);
                Assert.Equal(expected[row][1], objects[row].Active);
                Assert.Equal(expected[row][2], objects[row].Name);
                Assert.Equal(expected[row][3], objects[row].Created);
            }
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData("")]
    [InlineData("4294967296")]
    [InlineData("-1")]
    [InlineData("long")]
    [InlineData("surrogate")]
    public void DataReader_XmlMetadataRejectsInvalidStylesWithOriginalCellReference(string style) {
        string path = CreateCompactFastPathWorkbook();
        try {
            if (style == "long") style = new string('0', 80) + "x";
            if (style == "surrogate") style = new string('0', 31) + "😀0";
            ReplaceXmlMetadataWorksheet(path, "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
                + $"<row r=\"2\"><c r=\"&#67;2\" t=\"n\" s=\"{style}\"><v>42</v></c></row></sheetData></worksheet>");
            var error = Assert.Throws<InvalidDataException>(() => {
                using var reader = ExcelDocument.OpenDataReader(path);
            });
            Assert.Contains("cell C2 contains an invalid cell style index", error.Message);
        } finally { File.Delete(path); }
    }

    private static void ReplaceXmlMetadataWorksheet(string path, string xml) =>
        ReplaceZipEntry(path, "xl/worksheets/sheet1.xml",
            Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(xml)).ToArray());

    private sealed class XmlMetadataRecord {
        public decimal Id { get; set; }
        public bool Active { get; set; }
        public string? Name { get; set; }
        public DateTime Created { get; set; }
    }
}
