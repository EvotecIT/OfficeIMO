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
    [InlineData(false, null)]
    [InlineData(false, "SharedString")]
    [InlineData(true, null)]
    [InlineData(true, "SharedString")]
    [InlineData(false, "Utf8")]
    [InlineData(false, "XmlControl")]
    [InlineData(false, "Noncharacter")]
    [InlineData(false, "NormalizedText")]
    public void OpenReader_CompressedImplicitRowsPreserveWindowValuesAndValidateLateCells(bool boundedWindow, string? invalidTail) {
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
                if (invalidTail == "SharedString" && row == count) xml.Append("<c t=\"s\"><v>2147483647</v></c>");
                else {
                    xml.Append("<c t=\"inlineStr\"><is><t>Label ").Append(row).Append(padding);
                    if (invalidTail == "NormalizedText" && row == count) xml.Append("β&amp;Co&#13;raw\r\nend");
                    xml.Append("</t></is></c>");
                }
                if (row == 1 && boundedWindow) xml.Append("<c><f>1</f><v>1</v></c>");
                else xml.Append("<c><v>").Append(row).Append("</v></c>");
                xml.Append("<c s=\"1\"><v>")
                    .Append(45_000 + row % 365).Append("</v></c></row>");
            }
            xml.Append("</sheetData></worksheet>");
            byte[] bytes = Encoding.UTF8.GetBytes(xml.ToString());
            if (invalidTail is "Utf8" or "XmlControl" or "Noncharacter") {
                int tail = Array.LastIndexOf(bytes, (byte)'x');
                Assert.True(tail >= 2);
                bytes[tail] = invalidTail == "Utf8" ? (byte)0xff : (byte)0;
                if (invalidTail == "Noncharacter") {
                    bytes[tail - 2] = 0xef;
                    bytes[tail - 1] = 0xbf;
                    bytes[tail] = 0xbe;
                }
            }
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", bytes);
            var options = new ExcelReaderOptions { UseCachedFormulaResult = false };
            if (boundedWindow) {
                options.MaxDataReaderBufferedCells = 12;
                options.MaxDataReaderChunkRows = 2;
            }
            using var reader = ExcelDocument.OpenReader(path, options);
            if (invalidTail != null && invalidTail != "NormalizedText") {
                if (invalidTail == "SharedString") Assert.Throws<InvalidDataException>(() => reader.Read());
                else Assert.Throws<System.Xml.XmlException>(() => reader.Read());
                return;
            }
            for (int row = 1; row <= count; row++) {
                Assert.True(reader.Read());
                string expected = "Label " + row + padding;
                if (invalidTail == "NormalizedText" && row == count) expected += "β&Co\rraw\nend";
                Assert.Equal(expected, reader.GetString(0));
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
