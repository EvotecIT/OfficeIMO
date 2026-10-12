using OfficeIMO.Data;
using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, false)]
    [InlineData(false, true, true)]
    [InlineData(true, false, false)]
    [InlineData(true, false, true)]
    [InlineData(true, true, false)]
    [InlineData(true, true, true)]
    public void OpenDataReader_IndexedNumericGettersPreserveDateSystemsAndGetterOrder(bool date1904, bool elapsed, bool numericFirst) {
        (string NumberXml, double Number, string DateXml, double Serial)[] rows = {
            ("1.25", 1.25, "45351.25", 45351.25),
            ("-0", BitConverter.Int64BitsToDouble(ParsedNegativeZeroBits), "61.25", 61.25),
            ("1.234567890123456", 1.234567890123456, "&#52;5352.5", 45352.5),
            ("2.5E0", 2.5, "45353.75", 45353.75),
        };
        string cells = string.Concat(rows.Select(row =>
            $"<row><c><v>{row.NumberXml}</v></c><c s=\"1\"><v>{row.DateXml}</v></c><c/></row>"))
            + "<row><c/><c s=\"1\"><v>1E100</v></c><c><v>7</v></c></row>";
        string path = CreateIndexedNumericAccessWorkbook(cells, date1904, elapsed);
        try {
            using var reader = ExcelDocument.OpenDataReader(path);
            foreach (var row in rows) {
                Assert.True(reader.Read());
                if (numericFirst) {
                    Assert.Equal(BitConverter.DoubleToInt64Bits(row.Number), BitConverter.DoubleToInt64Bits(reader.GetDouble(0)));
                    Assert.Equal(row.Serial, reader.GetDouble(1));
                } else {
                    Assert.Equal(row.Number, Assert.IsType<double>(reader.GetValue(0)));
                }
                DateTime expectedDate = elapsed ? DateTime.FromOADate(row.Serial)
                    : ExcelDateSystemConverter.FromSerial(row.Serial, date1904 ? ExcelDateSystem.NineteenFour : ExcelDateSystem.NineteenHundred);
                Assert.Equal(expectedDate, reader.GetDateTime(1));
                Assert.Equal(expectedDate, Assert.IsType<DateTime>(reader.GetValue(1)));
                Assert.Equal(row.Serial, reader.GetDouble(1));
                Assert.Equal(BitConverter.DoubleToInt64Bits(row.Number), BitConverter.DoubleToInt64Bits(reader.GetDouble(0)));
                Assert.Equal(BitConverter.DoubleToInt64Bits(row.Number), BitConverter.DoubleToInt64Bits(Assert.IsType<double>(reader.GetValue(0))));
                Assert.True(reader.IsDBNull(2));
                var values = new object[3];
                Assert.Equal(3, reader.GetValues(values));
                Assert.Equal(expectedDate, values[1]);
                Assert.Equal(DBNull.Value, values[2]);
            }
            Assert.True(reader.Read());
            Assert.True(reader.IsDBNull(0));
            Assert.Throws<InvalidCastException>(() => reader.GetDouble(0));
            Assert.Equal(1E100, reader.GetDouble(1));
            Assert.ThrowsAny<ArgumentException>(() => reader.GetDateTime(1));
            Assert.Equal(1E100, reader.GetDouble(1));
            Assert.Equal(7, reader.GetInt32(2));
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_IndexedNumericGettersRetainFormulaDecimalAndConverterValues() {
        string path = CreateIndexedNumericAccessWorkbook(
            "<row><c><f>1+0.25</f><v>1.25</v></c><c s=\"1\"><f>45351+0.25</f><v>45351.25</v></c><c/></row>",
            date1904: false, elapsed: false);
        try {
            using (var reader = ExcelDocument.OpenDataReader(path)) {
                Assert.True(reader.Read());
                Assert.Equal(1.25d, reader.GetDouble(0));
                Assert.Equal(new DateTime(2024, 2, 29, 6, 0, 0), reader.GetDateTime(1));
                Assert.Equal(45351.25d, reader.GetDouble(1));
            }
            using (var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions { UseCachedFormulaResult = false })) {
                Assert.True(reader.Read());
                Assert.Equal("1+0.25", reader.GetValue(0));
                Assert.Throws<FormatException>(() => reader.GetDouble(0));
                Assert.Equal("45351+0.25", reader.GetValue(1));
            }
            using (var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions { NumericAsDecimal = true })) {
                Assert.True(reader.Read());
                Assert.Equal(1.25m, reader.GetDecimal(0));
                Assert.Equal(1.25m, Assert.IsType<decimal>(reader.GetValue(0)));
            }
            using (var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions {
                CellValueConverter = context => context.RawText == "1.25" ? new ExcelCellValue(9.75d) : ExcelCellValue.NotHandled
            })) {
                Assert.True(reader.Read());
                Assert.Equal(9.75d, reader.GetDouble(0));
                Assert.Equal(9.75d, Assert.IsType<double>(reader.GetValue(0)));
            }
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OpenDataReader_CoreNumericMappingPreservesDateAndSerialFromTheSameColumn(bool numericFirst) {
        string path = CreateIndexedNumericAccessWorkbook(
            "<row><c><v>1.25</v></c><c s=\"1\"><v>45351.25</v></c><c/></row>",
            date1904: true, elapsed: false);
        try {
            using var reader = ExcelDocument.OpenDataReader(path);
            var mapped = Assert.Single(reader.RowsAs<IndexedNumericMappingRecord>(mapper => {
                void Serial() => mapper.FromColumn<double>("Date", (row, value) => { row.Serial = value; return row; });
                void Date() => mapper.FromColumn<DateTime>("Date", (row, value) => { row.Date = value; return row; });
                if (numericFirst) { Serial(); Date(); } else { Date(); Serial(); }
            }));
            Assert.Equal(45351.25d, mapped.Serial);
            Assert.Equal(new DateTime(2028, 3, 1, 6, 0, 0), mapped.Date);
        } finally {
            File.Delete(path);
        }
    }

    private sealed class IndexedNumericMappingRecord {
        public double Serial { get; set; }
        public DateTime Date { get; set; }
    }

    private static string CreateIndexedNumericAccessWorkbook(string rows, bool date1904, bool elapsed) {
        string path = CreateCompactFastPathWorkbook();
        using (var document = ExcelDocument.Load(path)) {
            document.DateSystem = date1904 ? ExcelDateSystem.NineteenFour : ExcelDateSystem.NineteenHundred;
            document.Save();
        }
        string styles = "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
            + $"<cellXfs count=\"2\"><xf numFmtId=\"0\"/><xf numFmtId=\"{(elapsed ? 46 : 14)}\"/></cellXfs></styleSheet>";
        ReplaceZipEntry(path, "xl/styles.xml", Encoding.UTF8.GetBytes(styles));
        string header = "<row><c t=\"inlineStr\"><is><t>Number</t></is></c>"
            + "<c t=\"inlineStr\"><is><t>Date</t></is></c><c t=\"inlineStr\"><is><t>Keep</t></is></c></row>";
        string xml = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
            + "<dimension ref=\"A1:C6\"/><sheetData>" + header + rows + "</sheetData></worksheet>";
        ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
        return path;
    }
}
