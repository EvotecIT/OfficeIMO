using System.Data.Common;
using System.Threading;
using Xunit;

namespace OfficeIMO.Excel.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("native", "duplicate")]
        [InlineData("native", "descending")]
        [InlineData("native", "descendingDuplicate")]
        [InlineData("sdk", "duplicate")]
        [InlineData("sdk", "descending")]
        [InlineData("sdk", "descendingDuplicate")]
        [InlineData("explicitXml", "duplicate")]
        [InlineData("explicitXml", "descending")]
        [InlineData("explicitXml", "descendingDuplicate")]
        public void DataReader_XmlCellOrderPreservesLastValuesAcrossGetterOrders(string surface, string order) {
            string path = CreateXmlCellOrderWorkbook(surface == "sdk");
            try {
                string cells = order switch {
                    "duplicate" => "<c r=\"A2\"><v>1</v></c><c r=\"A2\"><v>2</v></c><c r=\"B2\"><v>3</v></c>",
                    "descending" => "<c r=\"B2\"><v>3</v></c><c r=\"A2\"><v>2</v></c>",
                    _ => "<c r=\"B2\"><v>1</v></c><c r=\"A2\"><v>2</v></c><c r=\"B2\"><v>3</v></c>"
                };
                string xml = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                    + "<dimension ref=\"A1:B4097\"/><sheetData>"
                    + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Id</t></is></c>"
                    + "<c r=\"B1\" t=\"inlineStr\"><is><t>Value</t></is></c></row>"
                    + "<row r=\"2\">" + cells + "</row>"
                    + "<row r=\"3\"><c r=\"A3\"><v>4</v></c><c r=\"B3\"><v>5</v></c></row>"
                    + "<row r=\"4\"><c r=\"A4\"><v>6</v></c></row>"
                    // A real later row selects the supported streaming range in both workbook surfaces.
                    + "<row r=\"4097\"><c r=\"A4097\"><v>7</v></c></row></sheetData></worksheet>";
                ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
                foreach (string firstAccess in new[] { "ascending", "descending", "bulk" }) {
                    var options = new ExcelReadOptions { MaxDataReaderBufferedCells = 2 };
                    using var owner = surface == "explicitXml" ? ExcelDocumentReader.Open(path, options) : null;
                    using DbDataReader reader = owner == null ? ExcelDocument.OpenDataReader(path, options)
                        : (DbDataReader)owner.GetSheet("Data").ReadRangeAsDataReader("A1:B4097", schemaSampleRows: 0);
                    DataReaderSchemaContractAssertions.AssertCanonicalSchema(reader);
                    Assert.Equal(2, reader.FieldCount);
                    Assert.Equal("Id", reader.GetName(0));
                    Assert.Equal("Value", reader.GetName(1));
                    Assert.True(reader.Read());
                    var values = new object[2];
                    if (firstAccess == "bulk") {
                        Assert.Equal(2, reader.GetValues(values));
                        Assert.Equal(new object[] { 2D, 3D }, values);
                    } else {
                        int ordinal = firstAccess == "ascending" ? 0 : 1;
                        Assert.Equal(ordinal == 0 ? 2 : 3, reader.GetInt32(ordinal));
                    }
                    for (int repeat = 0; repeat < 2; repeat++) {
                        Assert.Equal(2, reader.GetInt32(0));
                        Assert.Equal(3, reader.GetInt32(1));
                        Assert.Equal(2D, Assert.IsType<double>(reader.GetValue(0)));
                        Assert.Equal(3D, Assert.IsType<double>(reader.GetValue(1)));
                    }
                    Assert.Equal(2, reader.GetValues(values));
                    Assert.Equal(new object[] { 2D, 3D }, values);
                    Assert.True(reader.Read());
                    Assert.Equal(5, reader.GetInt32(1));
                    Assert.Equal(4, reader.GetInt32(0));
                    Assert.True(reader.Read());
                    Assert.True(reader.IsDBNull(1));
                    Assert.Equal(6, reader.GetInt32(0));
                    Assert.True(reader.Read());
                    Assert.True(reader.IsDBNull(0));
                    Assert.True(reader.IsDBNull(1));
                    reader.Close();
                    Assert.True(reader.IsClosed);
                }
            } finally {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData("native", false)]
        [InlineData("native", true)]
        [InlineData("sdk", false)]
        [InlineData("sdk", true)]
        [InlineData("explicitXml", false)]
        [InlineData("explicitXml", true)]
        public void DataReader_XmlCellOrderPreservesLastKindsNullsAndDateSerials(string surface, bool date1904) {
            string path = CreateXmlCellOrderWorkbook(surface == "sdk", date1904);
            try {
                const double serial = 45351.25D;
                DateTime expectedDate = ExcelDateSystemConverter.FromSerial(serial,
                    date1904 ? ExcelDateSystem.NineteenFour : ExcelDateSystem.NineteenHundred);
                string styles = "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                    + "<cellXfs count=\"2\"><xf numFmtId=\"0\"/><xf numFmtId=\"14\"/></cellXfs></styleSheet>";
                ReplaceZipEntry(path, "xl/styles.xml", Encoding.UTF8.GetBytes(styles));
                string xml = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                    + "<dimension ref=\"A1:F4097\"/><sheetData><row r=\"1\">"
                    + string.Concat(new[] { "Amount", "Missing", "Date", "Cleared", "Flag", "Text" }.Select((name, column) =>
                        $"<c r=\"{(char)('A' + column)}1\" t=\"inlineStr\"><is><t>{name}</t></is></c>"))
                    + "</row><row r=\"2\">"
                    + "<c r=\"A2\" s=\"1\"><v>1</v></c><c r=\"A2\"><v>12.5</v></c>"
                    + "<c r=\"C2\"><v>1</v></c><c r=\"C2\" s=\"1\"><v>45351.25</v></c>"
                    + "<c r=\"D2\" s=\"1\"><v>1</v></c><c r=\"D2\"/>"
                    + "<c r=\"E2\" t=\"inlineStr\"><is><t>old</t></is></c><c r=\"E2\" t=\"b\"><v>1</v></c>"
                    + "<c r=\"F2\"><v>1</v></c><c r=\"F2\" t=\"inlineStr\"><is/></c>"
                    + "</row><row r=\"4097\"><c r=\"A4097\"><v>7</v></c></row></sheetData></worksheet>";
                ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
                foreach (string firstAccess in new[] { "numeric", "decimal", "date", "null", "boolean", "bulk" }) {
                    var options = new ExcelReadOptions {
                        NumericAsDecimal = true, TreatDatesUsingNumberFormat = true, MaxDataReaderBufferedCells = 6
                    };
                    using var owner = surface == "explicitXml" ? ExcelDocumentReader.Open(path, options) : null;
                    using DbDataReader reader = owner == null ? ExcelDocument.OpenDataReader(path, options)
                        : (DbDataReader)owner.GetSheet("Data").ReadRangeAsDataReader("A1:F4097", schemaSampleRows: 0);
                    DataReaderSchemaContractAssertions.AssertCanonicalSchema(reader);
                    Assert.Equal(6, reader.FieldCount);
                    Assert.True(reader.Read());
                    var values = new object[6];
                    if (firstAccess == "numeric") Assert.Equal(serial, reader.GetDouble(2));
                    else if (firstAccess == "decimal") Assert.Equal(12.5M, reader.GetDecimal(0));
                    else if (firstAccess == "date") Assert.Equal(expectedDate, reader.GetDateTime(2));
                    else if (firstAccess == "null") Assert.True(reader.IsDBNull(3));
                    else if (firstAccess == "boolean") Assert.True(reader.GetBoolean(4));
                    else Assert.Equal(6, reader.GetValues(values));
                    Assert.Equal(12.5M, reader.GetDecimal(0));
                    Assert.Equal(12.5M, Assert.IsType<decimal>(reader.GetValue(0)));
                    Assert.True(reader.IsDBNull(1));
                    Assert.True(reader.IsDBNull(3));
                    Assert.True(reader.GetBoolean(4));
                    Assert.Equal(string.Empty, reader.GetString(5));
                    for (int repeat = 0; repeat < 2; repeat++) {
                        Assert.Equal(expectedDate, reader.GetDateTime(2));
                        Assert.Equal(expectedDate, Assert.IsType<DateTime>(reader.GetValue(2)));
                        Assert.Equal(serial, reader.GetDouble(2));
                    }
                    Assert.Equal(6, reader.GetValues(values));
                    Assert.Equal(new object[] { 12.5M, DBNull.Value, expectedDate, DBNull.Value, true, string.Empty }, values);
                }
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void DataReader_XmlCellOrderCancellationDoesNotPublishPartlyLoadedValues() {
            string path = CreateCompactFastPathWorkbook();
            try {
                const string xml = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                    + "<dimension ref=\"A1:B2\"/><sheetData>"
                    + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Id</t></is></c>"
                    + "<c r=\"B1\" t=\"inlineStr\"><is><t>Value</t></is></c></row>"
                    + "<row r=\"2\"><c r=\"A2\"><v>1</v></c><c r=\"A2\"><v>2</v></c>"
                    + "<c r=\"B2\"><v>3</v></c></row></sheetData></worksheet>";
                ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
                using var cancellation = new CancellationTokenSource();
                using var owner = ExcelDocumentReader.Open(path, new ExcelReadOptions {
                    CellValueConverter = context => {
                        if (context.RawText == "3") cancellation.Cancel();
                        return ExcelCellValue.NotHandled;
                    }
                });
                using var reader = owner.GetSheet("Data").ReadRangeAsDataReader("A1:B4097",
                    schemaSampleRows: 0, ct: cancellation.Token);
                Assert.True(reader.Read());
                Assert.ThrowsAny<OperationCanceledException>(() => reader.GetInt32(0));
                Assert.ThrowsAny<OperationCanceledException>(() => reader.GetValue(0));
                Assert.ThrowsAny<OperationCanceledException>(() => reader.IsDBNull(0));
                Assert.ThrowsAny<OperationCanceledException>(() => reader.GetValues(new object[2]));
                reader.Close();
                Assert.True(reader.IsClosed);
            } finally {
                File.Delete(path);
            }
        }

        private static string CreateXmlCellOrderWorkbook(bool multipleSheets, bool date1904 = false) {
            string path = CreateCompactFastPathWorkbook();
            using var document = ExcelDocument.Load(path);
            document.DateSystem = date1904 ? ExcelDateSystem.NineteenFour : ExcelDateSystem.NineteenHundred;
            if (multipleSheets) document.AddWorksheet("Other").CellValue(1, 1, "Other sheet");
            document.Save();
            return path;
        }
    }
}
