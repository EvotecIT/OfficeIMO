using OfficeIMO.Data;
using System.Globalization;
using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData("inline", false)]
    [InlineData("inline", true)]
    [InlineData("string", false)]
    [InlineData("string", true)]
    [InlineData("shared", false)]
    [InlineData("shared", true)]
    [InlineData("utf16", false)]
    [InlineData("utf16", true)]
    public void OpenDataReader_IndexedStringKeepsCacheOrderAndNormalizedValuesAcrossWindows(string storage, bool valueFirst) {
        (string Xml, string Text)[] values = {
            ("alpha", "alpha"), ("Zażółć 😀", "Zażółć 😀"),
            ("A&amp;B &#xD;", "A&B \r"), ("line\r\nbreak", "line\nbreak"),
            (string.Empty, string.Empty), ("123,5", "123,5"),
        };
        string rows = string.Concat(values.Select((value, index) => {
            string cell = storage switch {
                "shared" => "<c t=\"s\"><v>" + index + "</v></c>",
                "string" => "<c t=\"str\"><v>" + value.Xml + "</v></c>",
                _ => "<c t=\"inlineStr\"><is><t xml:space=\"preserve\">" + value.Xml + "</t></is></c>",
            };
            return "<row>" + cell + "<c><v>" + (index + 1) + "</v></c></row>";
        }));
        string xml = WindowedIndexWorksheet(rows, "A1:B6");
        string path = CreateWindowedIndexWorkbook(xml);
        try {
            if (storage == "shared") {
                string shared = "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"6\" uniqueCount=\"6\">"
                    + string.Concat(values.Select(value => "<si><t xml:space=\"preserve\">" + value.Xml + "</t></si>")) + "</sst>";
                ReplaceZipEntry(path, "xl/sharedStrings.xml", Encoding.UTF8.GetBytes(shared));
            }
            if (storage == "utf16") {
                byte[] bytes = Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(xml)).ToArray();
                ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", bytes);
            }
            using var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions {
                HasHeaderRow = false, MaxDataReaderBufferedCells = 2, MaxDataReaderChunkRows = 1,
                Culture = CultureInfo.GetCultureInfo("fr-FR"),
            });
            for (int row = 0; row < values.Length; row++) {
                Assert.True(reader.Read());
                if (valueFirst) Assert.Equal(values[row].Text, reader.GetValue(0));
                Assert.Equal(values[row].Text, reader.GetString(0));
                Assert.Equal(values[row].Text, Assert.IsType<string>(reader.GetValue(0)));
                Assert.Equal(values[row].Text, reader.GetFieldValue<string>(0));
                Assert.Equal(row + 1, reader.GetInt32(1));
                Assert.False(reader.IsDBNull(0));
                var current = new object[2];
                Assert.Equal(2, reader.GetValues(current));
                Assert.Equal(values[row].Text, current[0]);
                var chars = new char[values[row].Text.Length];
                Assert.Equal(chars.Length, reader.GetChars(0, 0, chars, 0, chars.Length));
                Assert.Equal(values[row].Text, new string(chars));
#if NET8_0_OR_GREATER
                if (reader.TryGetUtf8Text(0, out var borrowed)) {
                    Assert.Equal(values[row].Text, Encoding.UTF8.GetString(borrowed));
                }
#endif
            }
            Assert.Equal(123.5, reader.GetDouble(0));
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OpenDataReader_IndexedStringRetainsNullAndNonstringConversion(bool useCachedFormula) {
        string rows = "<row><c t=\"inlineStr\"><is><t/></is></c><c><v>1.25</v></c><c s=\"1\"><v>45292.25</v></c><c t=\"b\"><v>1</v></c><c><f>SUM(1,2)</f><v>3</v></c><c t=\"e\"><v>#N/A</v></c><c t=\"str\"><f>CONCAT(\"A\",\"B\")</f><v>A&amp;B</v></c></row>"
            + "<row><c/><c><v>2.5</v></c><c s=\"1\"><v>45293</v></c><c t=\"b\"><v>0</v></c><c><f>SUM(2,2)</f><v>4</v></c><c t=\"e\"><v>#VALUE!</v></c><c t=\"str\"><f>CONCAT(\"C\",\"D\")</f><v>CD</v></c></row>";
        string path = CreateWindowedIndexWorkbook(WindowedIndexWorksheet(rows, "A1:G2"));
        var culture = CultureInfo.GetCultureInfo("fr-FR");
        try {
            using var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions {
                HasHeaderRow = false, MaxDataReaderBufferedCells = 7, MaxDataReaderChunkRows = 1,
                Culture = culture, UseCachedFormulaResult = useCachedFormula,
            });
            Assert.Throws<InvalidOperationException>(() => reader.GetString(-1));
            Assert.True(reader.Read());
            Assert.Throws<IndexOutOfRangeException>(() => reader.GetString(-1));
            Assert.Throws<IndexOutOfRangeException>(() => reader.GetString(reader.FieldCount));
            Assert.Equal(string.Empty, reader.GetString(0));
            Assert.Equal("1,25", reader.GetString(1));
            Assert.Equal(45292.25, reader.GetDouble(2));
            Assert.Equal(DateTime.FromOADate(45292.25).ToString(culture), reader.GetString(2));
            Assert.Equal("True", reader.GetString(3));
            Assert.Equal(useCachedFormula ? "3" : "SUM(1,2)", reader.GetString(4));
            Assert.Equal("#N/A", reader.GetString(5));
            Assert.Equal(useCachedFormula ? "A&B" : "CONCAT(\"A\",\"B\")", reader.GetString(6));
            Assert.True(reader.Read());
            Assert.Throws<InvalidCastException>(() => reader.GetString(0));
            Assert.Equal(DBNull.Value, reader.GetValue(0));
            Assert.Equal("2,5", reader.GetString(1));
            Assert.Equal("False", reader.GetString(3));
            Assert.Equal(useCachedFormula ? "4" : "SUM(2,2)", reader.GetString(4));
            Assert.Equal(useCachedFormula ? "CD" : "CONCAT(\"C\",\"D\")", reader.GetString(6));
            Assert.False(reader.Read());
            Assert.Throws<InvalidOperationException>(() => reader.GetString(0));
            reader.Close();
            Assert.Throws<InvalidOperationException>(() => reader.GetString(-1));
        } finally {
            File.Delete(path);
        }
    }
}
