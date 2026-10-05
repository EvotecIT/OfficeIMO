using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Reader_Utf16PrefixDeclinePreservesTypedAndTabularValues(bool bigEndian, bool bom) {
        string path = CreateCompactFastPathWorkbook();
        var encoding = new UnicodeEncoding(bigEndian, bom);
        const string text = "Żółć 漢字";
        string xml = $$"""
            <?xml version="1.0" encoding="utf-16"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B3"/><sheetData>
              <row r="1"><c r="A1" t="str"><v>Id</v></c><c r="B1" t="str"><v>Name</v></c></row>
              <row r="2"><c r="A2"><v>42</v></c><c r="B2" t="str"><v>{{text}}</v></c></row>
              <row r="3"><c r="A3"><v>43</v></c><c r="B3" t="str"><v/></c></row>
            </sheetData></worksheet>
            """;
        try {
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", encoding.GetPreamble().Concat(encoding.GetBytes(xml)).ToArray());
            using (var owner = ExcelDocumentReader.Open(path)) {
                var sheet = owner.GetSheet("Data");
                foreach (var rows in new[] {
                    sheet.ReadObjects<Utf16PrefixRow>("A1:B3", ExcelExecutionMode.Sequential).ToArray(),
                    sheet.ReadObjectsStream<Utf16PrefixRow>("A1:B3").ToArray()
                }) {
                    Assert.Equal(2, rows.Length);
                    Assert.Equal(42, rows[0].Id);
                    Assert.Equal(text, rows[0].Name);
                    Assert.Equal(43, rows[1].Id);
                    Assert.Equal(string.Empty, rows[1].Name);
                }
                using var range = sheet.ReadRangeAsDataReader("A1:B3", schemaSampleRows: 0);
                Assert.True(range.Read());
                Assert.Equal(42, range.GetInt32(0));
                Assert.Equal(text, range.GetString(1));
                Assert.True(range.Read());
                Assert.Equal(43, range.GetInt32(0));
                Assert.Equal(string.Empty, range.GetString(1));
                Assert.False(range.Read());
            }
            using var reader = ExcelDocument.OpenDataReader(path);
            Assert.Equal(2, reader.FieldCount);
            Assert.Equal("Id", reader.GetName(0));
            Assert.Equal("Name", reader.GetName(1));
            Assert.True(reader.Read());
            Assert.Equal(42, reader.GetInt32(0));
            Assert.Equal(text, reader.GetString(1));
            Assert.True(reader.Read());
            Assert.Equal(43, reader.GetInt32(0));
            Assert.Equal(string.Empty, reader.GetString(1));
            Assert.False(reader.Read());
        } finally { File.Delete(path); }
    }

    public sealed class Utf16PrefixRow {
        public int Id { get; set; }
        public string? Name { get; set; }
    }
}
