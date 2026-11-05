using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OpenDataReader_Utf8InlineTextNormalizesLiteralLineEndingsBeforeReferences(bool explicitCoordinates) {
        (string Xml, string Expected)[] values = {
            ("first\r\nsecond\rthird", "first\nsecond\nthird"),
            ("a\r\r\nb\n\rc", "a\n\nb\n\nc"),
            ("trailing\r", "trailing\n"),
            ("kept&#13;\r\nreference", "kept\r\nreference"),
            ("&amp;\r\n&lt;tag&gt;", "&\n<tag>"),
            ("Zażółć\r\n\U0001F600\r", "Zażółć\n\U0001F600\n"),
        };
        string path = CreateCompactFastPathWorkbook();
        try {
            var rows = new StringBuilder();
            for (int index = 0; index < values.Length; index++) {
                int row = index + 2;
                rows.Append(explicitCoordinates
                    ? $"<row r=\"{row}\"><c r=\"A{row}\" t=\"inlineStr\"><is><t>{values[index].Xml}</t></is></c><c r=\"B{row}\"><v>{index}</v></c></row>"
                    : $"<row><c t=\"inlineStr\"><is><t>{values[index].Xml}</t></is></c><c><v>{index}</v></c></row>");
            }
            string header = explicitCoordinates
                ? "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Text</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>Index</t></is></c></row>"
                : "<row><c t=\"inlineStr\"><is><t>Text</t></is></c><c t=\"inlineStr\"><is><t>Index</t></is></c></row>";
            string xml = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                + $"<dimension ref=\"A1:B{values.Length + 1}\"/><sheetData>" + header + rows + "</sheetData></worksheet>";
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));

            using var reader = ExcelDocument.OpenReader(path);
            for (int index = 0; index < values.Length; index++) {
                Assert.True(reader.Read());
                Assert.Equal(index, reader.GetInt32(1));
                Assert.Equal(values[index].Expected, reader.GetString(0));
            }
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }
}
