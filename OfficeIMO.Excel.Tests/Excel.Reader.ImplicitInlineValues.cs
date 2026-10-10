using System.Text;
using System.Xml;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OpenDataReader_ImplicitCoordinatesPreserveInlineText(bool includeDimension) {
        string path = CreateCompactFastPathWorkbook();
        try {
            (string Xml, string Expected)[] values = {
                ("<t>alpha</t>", "alpha"),
                ("<t>Zażółć 😀</t>", "Zażółć 😀"),
                ("<t>AT&amp;T &lt;tag&gt; &#x1F600;</t>", "AT&T <tag> 😀"),
                ("<t>first\r\nsecond\rthird</t>", "first\nsecond\nthird"),
                ("<t xml:space=\"preserve\"> leading and trailing </t>", " leading and trailing "),
                ("<r><t>first</t></r><r><t> second</t></r>", "first second"),
                ("<t>al<![CDATA[pha]]>beta</t>", "alphabeta"),
                ("<t></t>", string.Empty),
                ("<t/>", string.Empty),
            };
            string dimension = includeDimension ? $"<dimension ref=\"A1:B{values.Length + 1}\"/>" : string.Empty;
            string rows = string.Concat(values.Select((value, index) =>
                $"<row><c t=\"inlineStr\"><is>{value.Xml}</is></c><c><v>{index + 1}</v></c></row>"));
            string xml = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                + dimension + "<sheetData><row><c t=\"inlineStr\"><is><t>Text</t></is></c>"
                + "<c t=\"inlineStr\"><is><t>Number</t></is></c></row>" + rows + "</sheetData></worksheet>";
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));

            using var reader = ExcelDocument.OpenDataReader(path);
            Assert.Equal(2, reader.FieldCount);
            Assert.Equal("Text", reader.GetName(0));
            Assert.Equal("Number", reader.GetName(1));
            for (int row = 0; row < values.Length; row++) {
                Assert.True(reader.Read());
                Assert.Equal(values[row].Expected, reader.GetString(0));
                Assert.Equal(row + 1, reader.GetInt32(1));
            }
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("<is><t>value</v></is></c>")]
    [InlineData("<is><t>value</t></v></c>")]
    [InlineData("<is><t>value</t></is></v>")]
    public void OpenDataReader_ImplicitInlineTextRejectsMalformedClosuresBeforeDelivery(string content) {
        string path = CreateCompactFastPathWorkbook();
        try {
            string xml = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                + "<dimension ref=\"A1:A2\"/><sheetData>"
                + "<row><c t=\"inlineStr\"><is><t>Text</t></is></c></row>"
                + "<row><c t=\"inlineStr\">" + content + "</row></sheetData></worksheet>";
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));

            Assert.Throws<XmlException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }
}
