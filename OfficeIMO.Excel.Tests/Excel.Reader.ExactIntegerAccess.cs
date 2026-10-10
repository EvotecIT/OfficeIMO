using OfficeIMO.Data;
using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    // Framework's canonical floating-point parsers normalize lexical negative zero;
    // modern .NET preserves its sign, including after an integer getter runs first.
#if NETFRAMEWORK
    private const long ParsedNegativeZeroBits = 0;
#else
    private const long ParsedNegativeZeroBits = long.MinValue;
#endif

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OpenDataReader_IntegerAccessPreservesRoundingAndGetterOrder(bool readDoubleFirst) {
        (string Raw, int Integer, double Number)[] cases = {
            ("0", 0, 0), ("42", 42, 42), ("+00042", 42, 42),
            ("-2147483648", int.MinValue, int.MinValue), ("2147483647", int.MaxValue, int.MaxValue),
            ("-0", 0, BitConverter.Int64BitsToDouble(ParsedNegativeZeroBits)),
            ("-000", 0, BitConverter.Int64BitsToDouble(ParsedNegativeZeroBits)),
            ("1.5", 2, 1.5), ("2.5", 2, 2.5), ("-1.5", -2, -1.5),
            ("4e1", 40, 40), ("1.25e1", 12, 12.5),
            ("2147483647.49", int.MaxValue, 2147483647.49), (" &#52;2 ", 42, 42)
        };
        string path = CreateExactIntegerAccessWorkbook(cases.Select(value => $"<c><v>{value.Raw}</v></c>"));
        try {
            using var reader = ExcelDocument.OpenDataReader(path);
            foreach (var item in cases) {
                Assert.True(reader.Read());
                if (readDoubleFirst) Assert.Equal(item.Number, reader.GetDouble(0));
                Assert.Equal(item.Integer, reader.GetInt32(0));
                Assert.Equal(BitConverter.DoubleToInt64Bits(item.Number), BitConverter.DoubleToInt64Bits(reader.GetDouble(0)));
                double materialized = Assert.IsType<double>(reader.GetValue(0));
                Assert.Equal(BitConverter.DoubleToInt64Bits(item.Number), BitConverter.DoubleToInt64Bits(materialized));
                Assert.Equal((long)item.Integer, reader.GetInt64(0));
            }
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("2147483648", 2147483648d)]
    [InlineData("1E100", 1E100)]
    public void OpenDataReader_IntegerOverflowPreservesTheNumericValue(string raw, double expected) {
        string path = CreateExactIntegerAccessWorkbook(new[] { $"<c><v>{raw}</v></c>" });
        try {
            using var reader = ExcelDocument.OpenDataReader(path);
            Assert.True(reader.Read());
            Assert.Throws<OverflowException>(() => reader.GetInt32(0));
            Assert.Equal(expected, reader.GetDouble(0));
            Assert.Equal(expected, Assert.IsType<double>(reader.GetValue(0)));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_CoreIntegerMappingPreservesLargeValueRounding() {
        string[] raw = { "-2147483648", "2147483647", "2147483648", "9007199254740993", "-9007199254740993", "1.5", "2.5", "-2.5" };
        long[] expected = { int.MinValue, int.MaxValue, 2147483648L, 9007199254740992L, -9007199254740992L, 2L, 2L, -2L };
        string path = CreateExactIntegerAccessWorkbook(raw.Select(value => $"<c><v>{value}</v></c>"));
        try {
            using var reader = ExcelDocument.OpenDataReader(path);
            Assert.Equal(expected, reader.RowsAs<ExactIntegerMappingRecord>().Select(row => row.Value).ToArray());
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OpenDataReader_IntegerAccessRetainsMissingFormulaAndDateValues(bool useCachedFormulaResult) {
        string path = CreateCompactFastPathWorkbook();
        try {
            ReplaceZipEntry(path, "xl/styles.xml", Encoding.UTF8.GetBytes("""
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs>
                </styleSheet>
                """));
            string xml = """
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B5"/><sheetData>
                <row><c t="inlineStr"><is><t>Value</t></is></c><c t="inlineStr"><is><t>Keep</t></is></c></row>
                <row><c><v>42</v></c><c><v>1</v></c></row>
                <row><c r="B3"><v>1</v></c></row>
                <row><c><f>40+2</f><v>42</v></c><c><v>1</v></c></row>
                <row><c s="1"><v>45292</v></c><c><v>1</v></c></row>
                </sheetData></worksheet>
                """;
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
            using var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions { UseCachedFormulaResult = useCachedFormulaResult });
            Assert.True(reader.Read());
            Assert.Equal(42, reader.GetInt32(0));
            Assert.True(reader.Read());
            Assert.True(reader.IsDBNull(0));
            Assert.Throws<InvalidCastException>(() => reader.GetInt32(0));
            Assert.True(reader.Read());
            if (useCachedFormulaResult) {
                Assert.Equal(42, reader.GetInt32(0));
                Assert.Equal(42d, Assert.IsType<double>(reader.GetValue(0)));
            } else {
                Assert.Throws<FormatException>(() => reader.GetInt32(0));
                Assert.Equal("40+2", reader.GetValue(0));
            }
            Assert.True(reader.Read());
            Assert.Equal(45292, reader.GetInt32(0));
            Assert.Equal(DateTime.FromOADate(45292), reader.GetDateTime(0));
            Assert.Equal(DateTime.FromOADate(45292), Assert.IsType<DateTime>(reader.GetValue(0)));
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_IntegerAccessHonorsDecimalAndCellConversionOptions() {
        string path = CreateExactIntegerAccessWorkbook(new[] { "<c><v>42</v></c>" });
        try {
            using (var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions { NumericAsDecimal = true })) {
                Assert.True(reader.Read());
                Assert.Equal(42, reader.GetInt32(0));
                Assert.Equal(42m, Assert.IsType<decimal>(reader.GetValue(0)));
            }
            using (var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions {
                CellValueConverter = context => context.RawText == "42" ? new ExcelCellValue(9.5d) : ExcelCellValue.NotHandled
            })) {
                Assert.True(reader.Read());
                Assert.Equal(10, reader.GetInt32(0));
                Assert.Equal(9.5d, reader.GetDouble(0));
            }
        } finally {
            File.Delete(path);
        }
    }

    private static string CreateExactIntegerAccessWorkbook(IEnumerable<string> cells) {
        string path = CreateCompactFastPathWorkbook();
        string[] rows = cells.ToArray();
        string xml = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
            + $"<dimension ref=\"A1:A{rows.Length + 1}\"/><sheetData>"
            + "<row><c t=\"inlineStr\"><is><t>Value</t></is></c></row>"
            + string.Concat(rows.Select(cell => "<row>" + cell + "</row>")) + "</sheetData></worksheet>";
        ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
        return path;
    }

    private sealed class ExactIntegerMappingRecord {
        public long Value { get; set; }
    }
}
