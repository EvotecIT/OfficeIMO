using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Fact]
    public void OpenReader_CompressedCdataRowClosingTextRetainsEntireCell() {
        string expected = CreateRowClosingMarkerText();
        string path = CreateCompactFastPathWorkbook();
        try {
            string xml = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
                + "<row><c t=\"inlineStr\"><is><t>Text</t></is></c></row>"
                + "<row><c t=\"inlineStr\"><is><t><![CDATA[" + expected
                + "]]></t></is></c></row></sheetData></worksheet>";
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
            using var reader = ExcelDocument.OpenReader(path);
            Assert.True(reader.Read());
            Assert.Equal(expected, reader.GetString(0));
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenReader_CompressedTrailingCommentRowClosingTextKeepsValidWorksheet() {
        string markers = CreateRowClosingMarkerText();
        string path = CreateCompactFastPathWorkbook();
        try {
            string xml = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
                + "<row><c t=\"inlineStr\"><is><t>Text</t></is></c></row>"
                + "<row><c t=\"inlineStr\"><is><t>Safe</t></is></c></row></sheetData><!--"
                + markers + markers + markers + "--></worksheet>";
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
            using var reader = ExcelDocument.OpenReader(path);
            Assert.True(reader.Read());
            Assert.Equal("Safe", reader.GetString(0));
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    private static string CreateRowClosingMarkerText() {
        var random = new Random(410);
        const string alphabet = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
        var text = new StringBuilder();
        for (int index = 0; index < 1_100_000; index++) {
            if (index % 4096 == 0) text.Append("</row>");
            text.Append(alphabet[random.Next(alphabet.Length)]);
        }
        return text.ToString();
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void OpenReader_CompressedImplicitRowsPreserveWindowValuesAndValidateLateCells(bool boundedWindow, bool invalidTail) {
        const int count = 6_000;
        string padding = new string('x', 320);
        string path = CreateCompactFastPathWorkbook();
        try {
            string styles = "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                + "<cellXfs count=\"2\"><xf numFmtId=\"0\"/><xf numFmtId=\"14\"/></cellXfs></styleSheet>";
            ReplaceZipEntry(path, "xl/styles.xml", Encoding.UTF8.GetBytes(styles));
            var xml = new StringBuilder("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row>");
            foreach (string header in new[] { "Label", "Id", "Date" })
                xml.Append("<c t=\"inlineStr\"><is><t>").Append(header).Append("</t></is></c>");
            xml.Append("</row>");
            for (int row = 1; row <= count; row++) {
                xml.Append("<row>");
                if (invalidTail && row == count) xml.Append("<c t=\"s\"><v>2147483647</v></c>");
                else xml.Append("<c t=\"inlineStr\"><is><t>Label ").Append(row).Append(padding).Append("</t></is></c>");
                if (row == 1) xml.Append("<c><f>1</f><v>1</v></c>");
                else xml.Append("<c><v>").Append(row).Append("</v></c>");
                xml.Append("<c s=\"1\"><v>")
                    .Append(45_000 + row % 365).Append("</v></c></row>");
            }
            xml.Append("</sheetData></worksheet>");
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml.ToString()));
            var options = new ExcelReaderOptions { UseCachedFormulaResult = false };
            if (boundedWindow) {
                options.MaxDataReaderBufferedCells = 12;
                options.MaxDataReaderChunkRows = 2;
            }
            using var reader = ExcelDocument.OpenReader(path, options);
            if (invalidTail) {
                Assert.Throws<InvalidDataException>(() => reader.Read());
                return;
            }
            for (int row = 1; row <= count; row++) {
                Assert.True(reader.Read());
                Assert.Equal("Label " + row + padding, reader.GetString(0));
                Assert.Equal(row, reader.GetInt32(1));
                double serial = 45_000 + row % 365;
                Assert.Equal(ExcelDateSystemConverter.FromSerial(serial, ExcelDateSystem.NineteenHundred), reader.GetDateTime(2));
                Assert.Equal(serial, reader.GetDouble(2));
            }
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }
}
