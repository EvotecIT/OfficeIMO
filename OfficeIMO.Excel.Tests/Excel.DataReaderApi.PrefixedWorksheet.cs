using System.Data;
using System.Text;
using System.Xml;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData("x", true)]
    [InlineData("sheet", true)]
    [InlineData("x", false)]
    [InlineData("", true)]
    public void DataReader_PrefixedWorksheetPreservesTypedValuesAndEmptyCells(string prefix, bool workbookReader) {
        string path = CreateCompactFastPathWorkbook();
        try {
            string xml = PrefixedWorksheetXml(prefix).Replace(" spaced", " literal\r\nline\r spaced");
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
            using var document = workbookReader ? null : ExcelDocumentReader.Open(path);
            using IDataReader reader = workbookReader
                ? ExcelDocument.OpenDataReader(path)
                : document!.GetSheet("Data").ReadRangeAsDataReader("A1:C3", schemaSampleRows: 0);
            Assert.Equal(3, reader.FieldCount);
            Assert.Equal("Value", reader.GetName(0));
            Assert.Equal("Text", reader.GetName(1));
            Assert.True(reader.Read());
            Assert.Equal(42, reader.GetInt32(0));
            Assert.Equal(" literal\nline\n spaced\r\n🌍 & <text> ", reader.GetString(1));
            Assert.True(reader.GetBoolean(2));
            Assert.True(reader.Read());
            Assert.Equal(43, reader.GetInt32(0));
            Assert.Equal(string.Empty, reader.GetString(1));
            Assert.True(reader.IsDBNull(2));
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void DataReader_PrefixedWorksheetRetainsFallbackForMixedElementPrefixes() {
        string path = CreateCompactFastPathWorkbook();
        try {
            string xml = PrefixedWorksheetXml("x").Replace("<x:v>43</x:v>",
                "<other:v xmlns:other=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">43</other:v>");
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
            using var reader = ExcelDocument.OpenDataReader(path);
            Assert.True(reader.Read());
            Assert.Equal(42, reader.GetInt32(0));
            Assert.True(reader.Read());
            Assert.Equal(43, reader.GetInt32(0));
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("row", "r=\"2\"")]
    [InlineData("c", "r=\"A2\"")]
    [InlineData("v", "")]
    [InlineData("f", "")]
    [InlineData("is", "")]
    [InlineData("t", "xml:space=\"preserve\"")]
    public void DataReader_PrefixedWorksheetRejectsReboundValueNamespaces(string element, string attributes) {
        string path = CreateCompactFastPathWorkbook();
        try {
            string xml = PrefixedWorksheetXml("x");
            string tag = "<x:" + element + (attributes.Length == 0 ? "" : " " + attributes) + ">";
            int start = xml.IndexOf(tag, StringComparison.Ordinal);
            Assert.True(start >= 0);
            xml = xml.Insert(start + tag.Length - 1, " xmlns:x=\"urn:foreign\"");
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
            Assert.Throws<InvalidDataException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void DataReader_PrefixedWorksheetStillValidatesMalformedTrailingXml() {
        string path = CreateCompactFastPathWorkbook();
        try {
            string xml = PrefixedWorksheetXml("x").Replace("</x:worksheet>", "<!--invalid--comment--></x:worksheet>");
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
            Assert.Throws<XmlException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }

    private static string PrefixedWorksheetXml(string prefix) {
        string xml = $$"""
        <?xml version="1.0" encoding="utf-8"?>
        <{{prefix}}:worksheet xmlns:{{prefix}}="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <{{prefix}}:dimension ref="A1:C3"/>
          <{{prefix}}:sheetData>
            <{{prefix}}:row r="1"><{{prefix}}:c r="A1" t="inlineStr"><{{prefix}}:is><{{prefix}}:t>Value</{{prefix}}:t></{{prefix}}:is></{{prefix}}:c><{{prefix}}:c r="B1" t="inlineStr"><{{prefix}}:is><{{prefix}}:t>Text</{{prefix}}:t></{{prefix}}:is></{{prefix}}:c><{{prefix}}:c r="C1" t="inlineStr"><{{prefix}}:is><{{prefix}}:t>Enabled</{{prefix}}:t></{{prefix}}:is></{{prefix}}:c></{{prefix}}:row>
            <{{prefix}}:row r="2"><{{prefix}}:c r="A2"><{{prefix}}:f>21*2</{{prefix}}:f><{{prefix}}:v>42</{{prefix}}:v></{{prefix}}:c><{{prefix}}:c r="B2" t="inlineStr"><{{prefix}}:is><{{prefix}}:t xml:space="preserve"> spaced&#xD;&#xA;🌍 &amp; &lt;text&gt; </{{prefix}}:t></{{prefix}}:is></{{prefix}}:c><{{prefix}}:c r="C2" t="b"><{{prefix}}:v>1</{{prefix}}:v></{{prefix}}:c></{{prefix}}:row>
            <{{prefix}}:row r="3"><{{prefix}}:c r="A3"><{{prefix}}:v>43</{{prefix}}:v></{{prefix}}:c><{{prefix}}:c r="B3" t="inlineStr"><{{prefix}}:is/></{{prefix}}:c></{{prefix}}:row>
          </{{prefix}}:sheetData>
        </{{prefix}}:worksheet>
        """;
        return prefix.Length == 0
            ? xml.Replace("<:", "<").Replace("</:", "</").Replace("xmlns:=", "xmlns=")
            : xml;
    }
}
