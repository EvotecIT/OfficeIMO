using OfficeIMO.Data;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Fact]
    public void OpenDataReader_DefaultBufferLimitAllowsMoreThanOneMillionTotalCells() {
        const int columns = 1000;
        const int rows = 1001;
        string path = CreateWindowedIndexWorkbook(WindowedIndexLargeWorksheet());
        try {
            using var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions { HasHeaderRow = false });
            Assert.Equal(columns, reader.FieldCount);
            for (int row = 1; row <= rows; row++) {
                Assert.True(reader.Read());
                Assert.Equal("row " + row, reader.GetString(0));
                Assert.True(reader.IsDBNull(columns - 1));
#if NET8_0_OR_GREATER
                Assert.True(reader.TryGetUtf8Text(0, out var text));
                Assert.Equal("row " + row, Encoding.UTF8.GetString(text));
#endif
            }
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OpenDataReader_OverDefaultCellLimitRejectsLateInvalidInputBeforeDelivery(bool malformedXml) {
        string xml = WindowedIndexLargeWorksheet();
        if (malformedXml) xml = xml.Replace("</worksheet>", "<broken></worksheet>");
        else {
            int lastRowEnd = xml.LastIndexOf("</row>", StringComparison.Ordinal);
            xml = xml.Remove(lastRowEnd - 4, 4).Insert(lastRowEnd - 4, "<c s=\"9\"/>");
        }
        string path = CreateWindowedIndexWorkbook(xml);
        try {
            if (malformedXml) Assert.Throws<XmlException>(() => ExcelDocument.OpenDataReader(path,
                new ExcelReadOptions { HasHeaderRow = false }));
            else Assert.Throws<InvalidDataException>(() => ExcelDocument.OpenDataReader(path,
                new ExcelReadOptions { HasHeaderRow = false }));
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void OpenDataReader_WindowBoundariesKeepTypedCachesAndNormalizedText(bool dimension, bool numericAsDecimal) {
        string rows = WindowedIndexHeader
            + "<row><c t=\"inlineStr\"><is><t>alpha</t></is></c><c><v>1</v></c><c s=\"1\"><v>45292.25</v></c><c><v>1.5</v></c></row>"
            + "<row><c t=\"inlineStr\"><is><t>β&amp;Co</t></is></c><c><v>2</v></c><c s=\"1\"><v>&#52;5293.5</v></c><c><v>-0</v></c></row>"
            + "<row><c t=\"inlineStr\"><is><t>line\r\nbreak</t></is></c><c><f>1+2</f><v>3</v></c><c s=\"1\"><v>45294.75</v></c><c><v>4.5</v></c></row>"
            + "<row><c/><c><v>4</v></c><c/><c><v>6</v></c></row>"
            + "<row><c t=\"inlineStr\"><is><t></t></is></c><c><v>5</v></c><c s=\"1\"><v>45295</v></c><c><v>7.5</v></c></row>";
        string path = CreateWindowedIndexWorkbook(WindowedIndexWorksheet(rows, dimension ? "A1:D6" : null));
        try {
            using var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions {
                MaxDataReaderBufferedCells = 8, MaxDataReaderChunkRows = 2, NumericAsDecimal = numericAsDecimal,
            });
            string?[] names = { "alpha", "β&Co", "line\nbreak", null, string.Empty };
            double?[] dates = { 45292.25, 45293.5, 45294.75, null, 45295 };
            double[] values = { 1.5, 0, 4.5, 6, 7.5 };
            for (int row = 0; row < names.Length; row++) {
                Assert.True(reader.Read());
                Assert.Equal(row + 1, reader.GetInt32(1));
                Assert.Equal(values[row], reader.GetDouble(3));
                if (row == 1 && !numericAsDecimal) {
                    Assert.Equal(ParsedNegativeZeroBits, BitConverter.DoubleToInt64Bits(reader.GetDouble(3)));
                }
                if (names[row] == null) Assert.True(reader.IsDBNull(0));
                else Assert.Equal(names[row], reader.GetString(0));
                if (dates[row] is double serial) {
                    Assert.Equal(DateTime.FromOADate(serial), reader.GetDateTime(2));
                    Assert.Equal(serial, reader.GetDouble(2));
                    Assert.IsType<DateTime>(reader.GetValue(2));
                } else {
                    Assert.True(reader.IsDBNull(2));
                }
                if (numericAsDecimal) Assert.IsType<decimal>(reader.GetValue(1));
                else Assert.IsType<double>(reader.GetValue(1));
                var current = new object[4];
                Assert.Equal(4, reader.GetValues(current));
                Assert.Equal(names[row] ?? (object)DBNull.Value, current[0]);
            }
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OpenDataReader_WindowedSparseRowsKeepOffsetRangesAndLogicalGaps(bool selectedRange) {
        string rows = "<row r=\"2\"><c r=\"B2\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"C2\" t=\"inlineStr\"><is><t>Count</t></is></c><c r=\"D2\" t=\"inlineStr\"><is><t>Value</t></is></c></row>"
            + "<row r=\"6\"><c r=\"B6\" t=\"inlineStr\"><is><t>first</t></is></c><c r=\"D6\"><v>6</v></c></row>"
            + "<row/>"
            + "<row><c r=\"B8\" t=\"inlineStr\"><is><t>last</t></is></c><c><v>8</v></c><c><v>9</v></c></row>";
        string path = CreateWindowedIndexWorkbook(WindowedIndexWorksheet(rows, "B2:D9"));
        try {
            using var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions {
                A1Range = selectedRange ? "C3:D8" : null,
                HasHeaderRow = !selectedRange,
                MaxDataReaderBufferedCells = selectedRange ? 2 : 3,
                MaxDataReaderChunkRows = 1,
            });
            Assert.Equal(selectedRange ? 2 : 3, reader.FieldCount);
            for (int row = 3; row <= 8; row++) {
                Assert.True(reader.Read());
                if (row is 6 or 8) {
                    int valueColumn = selectedRange ? 1 : 2;
                    Assert.Equal(row == 6 ? 6 : 9, reader.GetInt32(valueColumn));
                    if (row == 6) Assert.True(reader.IsDBNull(selectedRange ? 0 : 1));
                    else Assert.Equal(8, reader.GetInt32(selectedRange ? 0 : 1));
                    if (!selectedRange) Assert.Equal(row == 6 ? "first" : "last", reader.GetString(0));
                } else {
                    for (int column = 0; column < reader.FieldCount; column++) Assert.True(reader.IsDBNull(column));
                }
            }
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("footer", typeof(XmlException))]
    [InlineData("style", typeof(InvalidDataException))]
    [InlineData("shared-string", typeof(InvalidDataException))]
    [InlineData("foreign-value", typeof(InvalidDataException))]
    [InlineData("shared-formula", typeof(NotSupportedException))]
    public void OpenDataReader_WindowedIndexRejectsLateInvalidInputBeforeDelivery(string failure, Type error) {
        string invalidCell = failure switch {
            "style" => "<c s=\"9\"><v>3</v></c>",
            "shared-string" => "<c t=\"s\"><v>999999</v></c>",
            "foreign-value" => "<c><v xmlns=\"urn:foreign\">3</v></c>",
            "shared-formula" => "<c><f t=\"shared\" si=\"0\"/><v>3</v></c>",
            _ => "<c><v>3</v></c>",
        };
        string rows = "<row><c><v>1</v></c></row><row><c><v>2</v></c></row><row>" + invalidCell + "</row>";
        string xml = WindowedIndexWorksheet(rows, "A1:A3");
        if (failure == "footer") xml = xml.Replace("</worksheet>", "<broken></worksheet>");
        string path = CreateWindowedIndexWorkbook(xml);
        try {
            Assert.Throws(error, () => ExcelDocument.OpenDataReader(path, new ExcelReadOptions {
                HasHeaderRow = false, MaxDataReaderBufferedCells = 1, MaxDataReaderChunkRows = 1,
                UseCachedFormulaResult = false,
            }));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_WindowRequiresCapacityForOneProjectedRow() {
        string path = CreateWindowedIndexWorkbook(WindowedIndexWorksheet(WindowedIndexHeader, "A1:D1"));
        try {
            var error = Assert.Throws<InvalidDataException>(() => ExcelDocument.OpenDataReader(path,
                new ExcelReadOptions { MaxDataReaderBufferedCells = 3 }));
            Assert.Contains(nameof(ExcelReadOptions.MaxDataReaderBufferedCells), error.Message);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task OpenDataReader_CanceledWindowAdvanceKeepsPreviousValuesAndCanBeRetried() {
        string rows = "<row><c t=\"inlineStr\"><is><t>first</t></is></c></row><row><c t=\"inlineStr\"><is><t>second</t></is></c></row>";
        string path = CreateWindowedIndexWorkbook(WindowedIndexWorksheet(rows, "A1:A2"));
        try {
            using var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions {
                HasHeaderRow = false, MaxDataReaderBufferedCells = 1, MaxDataReaderChunkRows = 1,
            });
            Assert.True(reader.Read());
            Assert.Equal("first", reader.GetString(0));
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => reader.ReadAsync(cancellation.Token));
            Assert.Equal("first", reader.GetString(0));
#if NET8_0_OR_GREATER
            Assert.True(reader.TryGetUtf8Text(0, out var text));
            Assert.Equal("first", Encoding.UTF8.GetString(text));
#endif
            Assert.True(reader.Read());
            Assert.Equal("second", reader.GetString(0));
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    private const string WindowedIndexHeader = "<row><c t=\"inlineStr\"><is><t>Name</t></is></c><c t=\"inlineStr\"><is><t>Id</t></is></c><c t=\"inlineStr\"><is><t>Date</t></is></c><c t=\"inlineStr\"><is><t>Value</t></is></c></row>";

    private static string WindowedIndexLargeWorksheet() {
        var xml = new StringBuilder("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><dimension ref=\"A1:ALL1001\"/><sheetData>");
        for (int row = 1; row <= 1001; row++) {
            xml.Append("<row><c t=\"inlineStr\"><is><t>row ").Append(row).Append("</t></is></c>");
            for (int column = 1; column < 1000; column++) xml.Append("<c/>");
            xml.Append("</row>");
        }
        return xml.Append("</sheetData></worksheet>").ToString();
    }

    private static string WindowedIndexWorksheet(string rows, string? dimension) =>
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        + (dimension == null ? string.Empty : "<dimension ref=\"" + dimension + "\"/>")
        + "<sheetData>" + rows + "</sheetData></worksheet>";

    private static string CreateWindowedIndexWorkbook(string worksheetXml) {
        string path = CreateCompactFastPathWorkbook();
        ReplaceZipEntry(path, "xl/styles.xml", Encoding.UTF8.GetBytes("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><cellXfs count=\"2\"><xf numFmtId=\"0\"/><xf numFmtId=\"14\"/></cellXfs></styleSheet>"));
        ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(worksheetXml));
        return path;
    }
}
