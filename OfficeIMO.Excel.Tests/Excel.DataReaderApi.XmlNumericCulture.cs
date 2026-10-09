using System.Data.Common;
using System.Globalization;
using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData("singleSheet", false)]
    [InlineData("singleSheet", true)]
    [InlineData("multiSheet", false)]
    [InlineData("multiSheet", true)]
    [InlineData("range", false)]
    [InlineData("range", true)]
    public void DataReader_XmlNumericWireValuesRemainInvariantUnderConfiguredCulture(string surface, bool numericAsDecimal) {
        string path = CreateXmlTextBudgetWorkbook(
            "<row r=\"2\"><c r=\"A2\"><v>1.5</v></c>"
            + "<c r=\"B2\" t=\"n\"><v>-2.75</v></c>"
            + "<c r=\"C2\" t=\"n\"><f>1+2.125</f><v>3.125</v></c>"
            + "<c r=\"D2\"><v>1.23456789e100</v></c>"
            + "<c r=\"E2\" t=\"inlineStr\"><is><t>1,5</t></is></c>"
            + "<c r=\"F2\" t=\"str\"><v>2,75</v></c></row>", columns: 6,
            multipleSheets: surface == "multiSheet");
        try {
            var options = new ExcelReadOptions {
                Culture = CultureInfo.GetCultureInfo("de-DE"),
                NumericAsDecimal = numericAsDecimal
            };
            using (var owner = surface == "range" ? ExcelDocumentReader.Open(path, options) : null)
            using (DbDataReader reader = owner == null ? ExcelDocument.OpenDataReader(path, options)
                : (DbDataReader)owner.GetSheet("Data").ReadRangeAsDataReader("A1:F4097", schemaSampleRows: 0)) {
                Assert.True(reader.Read());
                AssertXmlNumericCultureValue(reader.GetValue(0), 1.5D, numericAsDecimal);
                Assert.Equal(-2.75D, reader.GetDouble(1));
                Assert.Equal(3.125M, reader.GetDecimal(2));
                Assert.Equal(1.23456789E100, Assert.IsType<double>(reader.GetValue(3)));
                Assert.Equal("1,5", reader.GetString(4));
                Assert.Equal(1.5M, reader.GetDecimal(4));
                Assert.Equal("2,75", reader.GetString(5));
                Assert.Equal(2.75D, reader.GetDouble(5));
                var values = new object[6];
                Assert.Equal(6, reader.GetValues(values));
                AssertXmlNumericCultureValue(values[1], -2.75D, numericAsDecimal);
                AssertXmlNumericCultureValue(values[2], 3.125D, numericAsDecimal);
                Assert.Equal(1.23456789E100, Assert.IsType<double>(values[3]));
            }

            options.UseCachedFormulaResult = false;
            using (var owner = surface == "range" ? ExcelDocumentReader.Open(path, options) : null)
            using (DbDataReader reader = owner == null ? ExcelDocument.OpenDataReader(path, options)
                : (DbDataReader)owner.GetSheet("Data").ReadRangeAsDataReader("A1:F4097", schemaSampleRows: 0)) {
                Assert.True(reader.Read());
                Assert.Equal("1+2.125", reader.GetString(2));
                Assert.Equal(1.5D, reader.GetDouble(0));
            }
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void DataReader_XmlNumericWireValuesPreserveConverterRawTextAndCulture() {
        string path = CreateXmlTextBudgetWorkbook(
            "<row r=\"2\"><c r=\"A2\" t=\"n\"><v>1.5</v></c>"
            + "<c r=\"B2\"><f>1+2.125</f><v>3.125</v></c>"
            + "<c r=\"C2\"><v>5.5</v></c>"
            + "<c r=\"D2\" t=\"inlineStr\"><is><t>1,5</t></is></c></row>", columns: 4);
        try {
            var seenRawText = new List<string>();
            using DbDataReader reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions {
                Culture = CultureInfo.GetCultureInfo("de-DE"),
                CellValueConverter = context => {
                    Assert.Equal("de-DE", context.Culture.Name);
                    if (context.RawText != "1.5" && context.RawText != "3.125") return ExcelCellValue.NotHandled;
                    seenRawText.Add(context.RawText);
                    return new ExcelCellValue(decimal.Parse(context.RawText, CultureInfo.InvariantCulture) * 2M);
                }
            });
            Assert.True(reader.Read());
            Assert.Equal(3M, Assert.IsType<decimal>(reader.GetValue(0)));
            Assert.Equal(6.25M, Assert.IsType<decimal>(reader.GetValue(1)));
            Assert.Equal(5.5D, Assert.IsType<double>(reader.GetValue(2)));
            Assert.Equal(1.5D, reader.GetDouble(3));
            Assert.Equal(new[] { "1.5", "3.125" }, seenRawText);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DataReader_IndexedXmlNumericEntitiesUseInvariantWireValues(bool numericAsDecimal) {
        string path = CreateCompactFastPathWorkbook();
        try {
            string worksheet = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                + "<dimension ref=\"A1:E2\"/><sheetData><row r=\"1\">"
                + "<c r=\"A1\" t=\"inlineStr\"><is><t>Number</t></is></c>"
                + "<c r=\"B1\" t=\"inlineStr\"><is><t>Formula</t></is></c>"
                + "<c r=\"C1\" t=\"inlineStr\"><is><t>Date</t></is></c>"
                + "<c r=\"D1\" t=\"inlineStr\"><is><t>Large</t></is></c>"
                + "<c r=\"E1\" t=\"inlineStr\"><is><t>Text</t></is></c></row>"
                + "<row r=\"2\"><c r=\"A2\"><v>1&#46;5</v></c>"
                + "<c r=\"B2\" t=\"n\"><f>1+2.125</f><v>3&#46;125</v></c>"
                + "<c r=\"C2\" s=\"1\"><v>45351&#46;25</v></c>"
                + "<c r=\"D2\"><v>1&#46;23456789e100</v></c>"
                + "<c r=\"E2\" t=\"inlineStr\"><is><t>indexed</t></is></c></row></sheetData></worksheet>";
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(worksheet));
            SetXmlNumericCultureDateStyle(path);
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(path,
                new ExcelReadOptions { NumericAsDecimal = numericAsDecimal });
            Assert.True(reader.Read());
#if NET8_0_OR_GREATER
            Assert.True(reader.TryGetUtf8Text(4, out var borrowed));
            Assert.Equal("indexed", Encoding.UTF8.GetString(borrowed));
#endif
            Assert.Equal(1.5D, reader.GetDouble(0));
            AssertXmlNumericCultureValue(reader.GetValue(0), 1.5D, numericAsDecimal);
            AssertXmlNumericCultureValue(reader.GetValue(1), 3.125D, numericAsDecimal);
            Assert.Equal(ExcelDateSystemConverter.FromSerial(45351.25D, ExcelDateSystem.NineteenHundred), reader.GetDateTime(2));
            Assert.Equal(45351.25D, reader.GetDouble(2));
            Assert.Equal(1.23456789E100, Assert.IsType<double>(reader.GetValue(3)));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void DataReader_XmlWireNumberConversionKeepsDateSerialsDeferred() {
        string path = CreateXmlTextBudgetWorkbook(
            "<row r=\"2\"><c r=\"A2\" s=\"1\"><v>45351.25</v></c>"
            + "<c r=\"B2\" s=\"1\"><f>1.23456789e100</f><v>1.23456789e100</v></c></row>", columns: 2);
        try {
            SetXmlNumericCultureDateStyle(path);
            using DbDataReader reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions {
                NumericAsDecimal = true, MaxXmlDataReaderBufferedCharacters = 64
            });
            Assert.True(reader.Read());
            Assert.Equal(1.23456789E100, reader.GetDouble(1));
            Assert.Equal(45351.25D, reader.GetDouble(0));
            DateTime date = ExcelDateSystemConverter.FromSerial(45351.25D, ExcelDateSystem.NineteenHundred);
            Assert.Equal(date, reader.GetDateTime(0));
            Assert.Equal(date, Assert.IsType<DateTime>(reader.GetValue(0)));
            Assert.Equal(45351.25D, reader.GetDouble(0));
            Assert.Equal(1.23456789E100, reader.GetDouble(1));
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Reader_XmlNumericBindingsSeparateWireValuesFromConfiguredStringCulture(bool streamed) {
        string[] names = { "Amount", "Count", "LongCount", "TextAmount", "LocalizedAmount", "NullableCount" };
        string headers = string.Concat(names.Select((name, column) =>
            $"<c r=\"{(char)('A' + column)}1\" t=\"inlineStr\"><is><t>{name}</t></is></c>"));
        string path = CreateXmlTextBudgetWorkbook(
            "<row r=\"2\"><c r=\"A2\" t=\"n\"><v>1.5</v></c>"
            + "<c r=\"B2\"><v>1.0</v></c>"
            + "<c r=\"C2\"><f>1+1</f><v>2.0</v></c>"
            + "<c r=\"D2\" t=\"inlineStr\"><is><t>1.5</t></is></c>"
            + "<c r=\"E2\" t=\"str\"><v>1,5</v></c>"
            + "<c r=\"F2\"><v>1.0</v></c></row>", columns: names.Length, headerCells: headers);
        try {
            using ExcelDocumentReader owner = ExcelDocumentReader.Open(path,
                new ExcelReadOptions { Culture = CultureInfo.GetCultureInfo("de-DE") });
            var sheet = owner.GetSheet("Data");
            XmlNumericCultureBindingRow row = Assert.Single(streamed
                ? sheet.ReadObjectsStream<XmlNumericCultureBindingRow>("A1:F2")
                : sheet.ReadObjects<XmlNumericCultureBindingRow>("A1:F2", ExcelExecutionMode.Sequential));
            Assert.Equal(1.5D, row.Amount);
            Assert.Equal(1, row.Count);
            Assert.Equal(2L, row.LongCount);
            Assert.Equal(15D, row.TextAmount);
            Assert.Equal(1.5D, row.LocalizedAmount);
            Assert.Equal(1, row.NullableCount);
        } finally {
            File.Delete(path);
        }
    }

    private sealed class XmlNumericCultureBindingRow {
        public double Amount { get; set; }
        public int Count { get; set; }
        public long LongCount { get; set; }
        public double TextAmount { get; set; }
        public double LocalizedAmount { get; set; }
        public int? NullableCount { get; set; }
    }

    private static void AssertXmlNumericCultureValue(object value, double expected, bool numericAsDecimal) {
        if (numericAsDecimal) Assert.Equal((decimal)expected, Assert.IsType<decimal>(value));
        else Assert.Equal(expected, Assert.IsType<double>(value));
    }

    private static void SetXmlNumericCultureDateStyle(string path) {
        ReplaceZipEntry(path, "xl/styles.xml", Encoding.UTF8.GetBytes(
            "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
            + "<cellXfs count=\"2\"><xf numFmtId=\"0\"/><xf numFmtId=\"14\"/></cellXfs></styleSheet>"));
    }
}
