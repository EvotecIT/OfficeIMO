using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Reader_XmlCoordinateAttributesPreserveAllReadShapes(bool prefixed) {
        string path = CreateCompactFastPathWorkbook();
        try {
            string prefix = prefixed ? "x:" : string.Empty;
            string declaration = prefixed ? "xmlns:x" : "xmlns";
            string paddedReference = new string(' ', 48) + "B2" + new string(' ', 48);
            string xml = $$"""
                <?xml version="1.0" encoding="utf-16"?>
                <{{prefix}}worksheet {{declaration}}="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><{{prefix}}sheetData>
                  <{{prefix}}row r="&#49;">
                    <{{prefix}}c r="&#65;1" t="inlineStr"><{{prefix}}is><{{prefix}}t>Id</{{prefix}}t></{{prefix}}is></{{prefix}}c>
                    <{{prefix}}c r="B1" t="inlineStr"><{{prefix}}is><{{prefix}}t>Name</{{prefix}}t></{{prefix}}is></{{prefix}}c>
                    <{{prefix}}c r="C1" t="inlineStr"><{{prefix}}is><{{prefix}}t>Amount</{{prefix}}t></{{prefix}}is></{{prefix}}c>
                  </{{prefix}}row>
                  <{{prefix}}row r="&#50;">
                    <{{prefix}}c r="A2" s="0"><{{prefix}}v>42</{{prefix}}v></{{prefix}}c>
                    <{{prefix}}c r="{{paddedReference}}" t="str"><{{prefix}}v>Hello</{{prefix}}v></{{prefix}}c>
                    <{{prefix}}c r="&#67;2"><{{prefix}}v>12.5</{{prefix}}v></{{prefix}}c>
                  </{{prefix}}row>
                  <{{prefix}}row r="3">
                    <{{prefix}}c><{{prefix}}v>43</{{prefix}}v></{{prefix}}c>
                    <{{prefix}}c r="" t="str"><{{prefix}}v>World</{{prefix}}v></{{prefix}}c>
                    <{{prefix}}c r="C3"><{{prefix}}v>13.5</{{prefix}}v></{{prefix}}c>
                  </{{prefix}}row>
                </{{prefix}}sheetData></{{prefix}}worksheet>
                """;
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml",
                Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(xml)).ToArray());
            var options = new ExcelReadOptions { NumericAsDecimal = true, InferDataTableColumnTypes = false };
            object[][] expected = { new object[] { 42m, "Hello", 12.5m }, new object[] { 43m, "World", 13.5m } };
            using (var reader = ExcelDocument.OpenDataReader(path, options)) {
                for (int row = 0; row < expected.Length; row++) {
                    Assert.True(reader.Read());
                    for (int column = 0; column < 3; column++) Assert.Equal(expected[row][column], reader.GetValue(column));
                }
                Assert.False(reader.Read());
            }
            using (var owner = ExcelDocumentReader.Open(path, options)) {
                var sheet = owner.GetSheet("Data");
                var values = sheet.ReadRange("A2:C3", ExcelExecutionMode.Sequential);
                using var table = sheet.ReadRangeAsDataTable("A1:C3", headersInFirstRow: true);
                var objects = sheet.ReadObjects<XmlCoordinateRecord>("A1:C3").ToArray();
                Assert.Equal(2, table.Rows.Count);
                Assert.Equal(2, objects.Length);
                for (int row = 0; row < expected.Length; row++) {
                    for (int column = 0; column < 3; column++) {
                        Assert.Equal(expected[row][column], values[row, column]);
                        Assert.Equal(expected[row][column], table.Rows[row][column]);
                    }
                    Assert.Equal(expected[row][0], objects[row].Id);
                    Assert.Equal(expected[row][1], objects[row].Name);
                    Assert.Equal(expected[row][2], objects[row].Amount);
                }
            }
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("empty")]
    [InlineData("entity")]
    [InlineData("long")]
    public void DataReader_XmlCoordinateDiagnosticsPreserveOriginalReference(string kind) {
        string path = CreateCompactFastPathWorkbook();
        try {
            string? encoded = kind switch {
                "missing" => null,
                "empty" => string.Empty,
                "entity" => "&#65;2",
                _ => "A" + new string('0', 80) + "2"
            };
            string expected = kind == "missing" ? "(unknown cell)" : kind == "entity" ? "A2" : encoded!;
            string attribute = encoded == null ? string.Empty : $" r=\"{encoded}\"";
            string xml = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
                + $"<row r=\"2\"><c{attribute} s=\"999999\"><v>42</v></c></row></sheetData></worksheet>";
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml",
                Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(xml)).ToArray());
            var exception = Assert.Throws<InvalidDataException>(() => {
                using var reader = ExcelDocument.OpenDataReader(path);
            });
            Assert.Contains($"cell {expected} references a missing cell style", exception.Message);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("shared-string")]
    [InlineData("formula")]
    [InlineData("nested-value")]
    public void DataReader_XmlCoordinateDiagnosticsSurviveReadingCellContent(string failure) {
        string path = CreateCompactFastPathWorkbook();
        try {
            string reference = "B2";
            string type = failure == "formula" ? string.Empty : " t=\"s\"";
            string content = failure switch {
                "shared-string" => "<v>999999</v>",
                "formula" => "<f t=\"shared\" si=\"0\"/>",
                _ => "<v><nested>0</nested></v>"
            };
            string xml = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
                + $"<row r=\"2\"><c r=\"A2\"><v>42</v></c><c r=\"{reference}\"{type}>{content}</c></row></sheetData></worksheet>";
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml",
                Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(xml)).ToArray());
            var exception = Record.Exception(() => {
                using var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions { UseCachedFormulaResult = false });
            });
            if (failure == "formula") {
                Assert.IsType<NotSupportedException>(exception);
                Assert.Contains($"'Data'!{reference}", exception!.Message);
            } else {
                Assert.IsType<InvalidDataException>(exception);
                Assert.Contains($"cell {reference}", exception!.Message);
            }
        } finally {
            File.Delete(path);
        }
    }

    private sealed class XmlCoordinateRecord {
        public decimal Id { get; set; }
        public string? Name { get; set; }
        public decimal Amount { get; set; }
    }
}
