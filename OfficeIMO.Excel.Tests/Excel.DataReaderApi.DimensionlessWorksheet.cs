using System.Text;
using System.Xml;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void DataReader_DimensionlessWorksheetPreservesSparseAndImplicitCoordinates(bool implicitRows, bool prefixed) {
        string path = CreateCompactFastPathWorkbook();
        try {
            string row4 = implicitRows ? string.Empty : " r=\"4\"";
            string xml = $$"""
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <sheetData>
                    <row r="2"/>
                    <row r="3"><c r="B3" t="inlineStr"><is><t>Id</t></is></c><c r="C3" t="inlineStr"><is><t>Value</t></is></c></row>
                    <row{{row4}}><c r="B4"><v>42</v></c><c><v>7</v></c></row>
                    <row r="6"><c r="B6"><v>43</v></c><c r="D6"><v>9</v></c></row>
                    <row r="10"/>
                  </sheetData>
                </worksheet>
                """;
            if (prefixed) {
                var document = System.Xml.Linq.XDocument.Parse(xml);
                document.Root!.SetAttributeValue(System.Xml.Linq.XNamespace.Xmlns + "s", document.Root.Name.NamespaceName);
                document.Root.Attribute("xmlns")!.Remove();
                xml = document.ToString();
            }
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
            using var reader = ExcelDocument.OpenDataReader(path);
            Assert.Equal(3, reader.FieldCount);
            Assert.Equal("Id", reader.GetName(0));
            Assert.Equal("Value", reader.GetName(1));
            Assert.True(reader.Read());
            Assert.Equal(42, reader.GetInt32(0));
            Assert.Equal(7, reader.GetInt32(1));
            Assert.True(reader.IsDBNull(2));
            Assert.True(reader.Read());
            Assert.True(reader.IsDBNull(0));
            Assert.True(reader.IsDBNull(1));
            Assert.True(reader.IsDBNull(2));
            Assert.True(reader.Read());
            Assert.Equal(43, reader.GetInt32(0));
            Assert.True(reader.IsDBNull(1));
            Assert.Equal(9, reader.GetInt32(2));
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("s=\"99999\"", "<v>42</v>", typeof(InvalidDataException))]
    [InlineData("t=\"s\"", "<v>99999</v>", typeof(InvalidDataException))]
    [InlineData("", "<f t=\"shared\" si=\"0\"/><v>42</v>", typeof(NotSupportedException))]
    public void DataReader_DimensionlessWorksheetValidatesBeforeExposingRows(string attributes, string content, Type error) {
        string path = CreateCompactFastPathWorkbook();
        try {
            string xml = $$"""
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
                  <row r="1"><c r="A1" t="inlineStr"><is><t>Value</t></is></c></row>
                  <row r="2"><c r="A2"><v>1</v></c></row>
                  <row r="3"><c r="A3" {{attributes}}>{{content}}</c></row>
                </sheetData></worksheet>
                """;
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
            Assert.Throws(error, () => ExcelDocument.OpenDataReader(path, new ExcelReadOptions { UseCachedFormulaResult = false }));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void DataReader_DimensionlessWorksheetValidatesTrailingXmlBeforeExposingRows() {
        string path = CreateCompactFastPathWorkbook();
        try {
            string xml = PrefixedWorksheetXml(string.Empty)
                .Replace("<dimension ref=\"A1:C3\"/>", string.Empty)
                .Replace("</worksheet>", "<broken></worksheet>");
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
            Assert.Throws<XmlException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }
}
