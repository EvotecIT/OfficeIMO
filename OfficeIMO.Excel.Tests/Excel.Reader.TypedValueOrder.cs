using System.Globalization;
using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData("date", false)]
    [InlineData("date", true)]
    [InlineData("decimal", false)]
    [InlineData("decimal", true)]
    [InlineData("double", false)]
    [InlineData("double", true)]
    [InlineData("boolean", false)]
    [InlineData("boolean", true)]
    public void Reader_PrimitiveNullChecksFollowBufferedRowTransitions(string kind, bool materialize) {
        string path = CreateCompactFastPathWorkbook();
        try {
            ReplaceZipEntry(path, "xl/styles.xml", Encoding.UTF8.GetBytes("""
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs>
                </styleSheet>
                """));
            string attributes = kind == "date" ? " s=\"1\"" : kind == "boolean" ? " t=\"b\"" : string.Empty;
            string xml = $$"""
                <?xml version="1.0" encoding="utf-16"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
                  <row r="1"><c r="A1"{{attributes}}><v>1</v></c></row>
                  <row r="3"><c r="A3"><v>3</v></c></row>
                  <row r="2"><c r="B2"><v>2</v></c></row>
                </sheetData></worksheet>
                """;
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(xml)).ToArray());
            using var owner = ExcelDocumentReader.Open(path, new ExcelReadOptions { NumericAsDecimal = kind == "decimal" });
            using var reader = owner.GetSheet("Data").ReadRangeAsDataReader("A1:B4098", headersInFirstRow: false, schemaSampleRows: 0);
            Assert.True(reader.Read());
            switch (kind) {
                case "date": Assert.Equal(new DateTime(1900, 1, 1), reader.GetDateTime(0)); break;
                case "decimal": Assert.Equal(1m, reader.GetDecimal(0)); break;
                case "double": Assert.Equal(1d, reader.GetDouble(0)); break;
                case "boolean": Assert.True(reader.GetBoolean(0)); break;
            }
            if (materialize) Assert.NotEqual(DBNull.Value, reader.GetValue(0));
            Assert.True(reader.Read());
            Assert.True(reader.IsDBNull(0));
            Assert.Equal(DBNull.Value, reader.GetValue(0));
            Assert.Equal(2, reader.GetInt32(1));
            Assert.True(reader.Read());
            Assert.Equal(3, reader.GetInt32(0));
            Assert.True(reader.IsDBNull(1));
            var values = new object[2];
            Assert.Equal(2, reader.GetValues(values));
            Assert.Equal(DBNull.Value, values[1]);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, true, false)]
    [InlineData(false, false, true)]
    [InlineData(true, false, true)]
    [InlineData(false, true, true)]
    [InlineData(true, true, true)]
    public void Reader_TypedGettersPreserveCanonicalValuesAcrossGetterOrder(bool utf16, bool date1904, bool streaming) {
        string path = CreateCompactFastPathWorkbook();
        try {
            using (var document = ExcelDocument.Load(path)) {
                document.DateSystem = date1904 ? ExcelDateSystem.NineteenFour : ExcelDateSystem.NineteenHundred;
                document.Save();
            }
            ReplaceZipEntry(path, "xl/styles.xml", Encoding.UTF8.GetBytes("""
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <cellXfs count="3"><xf numFmtId="0"/><xf numFmtId="14"/><xf numFmtId="46"/></cellXfs>
                </styleSheet>
                """));
            string xml = $$"""
                <?xml version="1.0" encoding="{{(utf16 ? "utf-16" : "utf-8")}}"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
                  <row r="1">
                    <c r="A1" s="1"><v>45351</v></c>
                    <c r="B1" s="1"><v>45352</v></c>
                    <c r="C1" s="2"><v>2.5</v></c>
                    <c r="D1"><v>165258.23999999999</v></c>
                    <c r="E1"><v>2.5</v></c>
                    <c r="F1" s="1"><v>1E100</v></c>
                  </row>
                  <row r="2"><c r="A2" s="1"><v>45353</v></c><c r="D2"><v>3.5</v></c></row>
                </sheetData></worksheet>
                """;
            Encoding encoding = utf16 ? Encoding.Unicode : new UTF8Encoding(false);
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", encoding.GetPreamble().Concat(encoding.GetBytes(xml)).ToArray());
            using var owner = ExcelDocumentReader.Open(path, new ExcelReadOptions {
                NumericAsDecimal = true,
                TreatDatesUsingNumberFormat = true,
                Culture = CultureInfo.InvariantCulture
            });
            using var reader = owner.GetSheet("Data").ReadRangeAsDataReader(streaming ? "A1:F4098" : "A1:F2", headersInFirstRow: false, schemaSampleRows: 0);
            var expectedDate = new DateTime(2024, 2, 29).AddDays(date1904 ? 1462 : 0);
            Assert.True(reader.Read());
            Assert.Equal(45351d, reader.GetDouble(0));
            Assert.Equal(expectedDate, reader.GetDateTime(0));
            Assert.Equal(expectedDate, Assert.IsType<DateTime>(reader.GetValue(0)));
            Assert.Equal(45351m, reader.GetDecimal(0));
            Assert.Equal(expectedDate.AddDays(1), reader.GetDateTime(1));
            Assert.Equal(45352, reader.GetInt32(1));
            Assert.Equal(expectedDate.AddDays(1), Assert.IsType<DateTime>(reader.GetValue(1)));
            Assert.Equal(DateTime.FromOADate(2.5), reader.GetDateTime(2));
            Assert.Equal(2.5d, reader.GetDouble(2));
            Assert.Equal(165258.24m, reader.GetDecimal(3));
            Assert.Equal(165258.24d, reader.GetDouble(3));
            Assert.Equal(165258.24m, Assert.IsType<decimal>(reader.GetValue(3)));
            Assert.Equal(2, reader.GetInt32(4));
            Assert.Equal((byte)2, reader.GetByte(4));
            Assert.Equal((short)2, reader.GetInt16(4));
            Assert.Equal(2L, reader.GetInt64(4));
            Assert.Equal(2.5f, reader.GetFloat(4));
            Assert.Equal(2.5m, Assert.IsType<decimal>(reader.GetValue(4)));
            Assert.Equal(1E100, reader.GetDouble(5));
            Assert.ThrowsAny<ArgumentException>(() => reader.GetDateTime(5));
            Assert.Equal(1E100, reader.GetDouble(5));

            Assert.True(reader.Read());
            Assert.Equal(expectedDate.AddDays(2), reader.GetDateTime(0));
            Assert.True(reader.IsDBNull(1));
            Assert.Equal(4, reader.GetInt32(3));
            var values = new object[6];
            Assert.Equal(6, reader.GetValues(values));
            Assert.Equal(expectedDate.AddDays(2), values[0]);
            Assert.Equal(DBNull.Value, values[1]);
            Assert.Equal(3.5m, Assert.IsType<decimal>(values[3]));
            int rowCount = 2;
            while (reader.Read()) {
                rowCount++;
                Assert.True(reader.IsDBNull(0));
                Assert.True(reader.IsDBNull(3));
            }
            Assert.Equal(streaming ? 4098 : 2, rowCount);
        } finally {
            File.Delete(path);
        }
    }
}
