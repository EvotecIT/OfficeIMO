using System.Text;
using System.Xml;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData("<c><v>42</v")]
    [InlineData("<c><v>42</v></c")]
    [InlineData("<c t=\"inlineStr\"><is><t>text</t></is")]
    [InlineData("<c t=\"inlineStr\"><is><t>text</t></is></c")]
    public void OpenDataReader_TruncatedCanonicalValueTagsRejectXmlBeforeDelivery(string truncatedCell) {
        string path = CreateCompactFastPathWorkbook();
        try {
            string xml = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                + "<dimension ref=\"A1:A2\"/><sheetData>"
                + "<row><c t=\"inlineStr\"><is><t>Value</t></is></c></row>"
                + "<row>" + truncatedCell;
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));

            Assert.Throws<XmlException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }
}
