using System.Data.Common;
using System.Text;
using System.Threading;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void XlsxTabularWorkbook_ReadsDimensionlessWorksheetWithoutSdkProjection(bool utf16) {
        string path = CreateCompactFastPathWorkbook();
        try {
            string encoding = utf16 ? "utf-16" : "utf-8";
            string xml = $$"""
                <?xml version="1.0" encoding="{{encoding}}"?>
                <s:worksheet xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <s:sheetData>
                    <s:row r="2"/>
                    <s:row r="3"><s:c r="B3" t="inlineStr"><s:is><s:t>Id</s:t></s:is></s:c><s:c r="C3" t="inlineStr"><s:is><s:t>Value</s:t></s:is></s:c></s:row>
                    <s:row><s:c r="B4"><s:v>42</s:v></s:c><s:c><s:v>7</s:v></s:c></s:row>
                    <s:row r="6"><s:c r="B6"><s:v>43</s:v></s:c><s:c r="D6"><s:v>9</s:v></s:c></s:row>
                    <s:row r="10"/>
                  </s:sheetData>
                </s:worksheet>
                """;
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml",
                (utf16 ? Encoding.Unicode : Encoding.UTF8).GetBytes(xml));
            using XlsxTabularWorkbook workbook = XlsxTabularWorkbook.Open(path, new ExcelReadOptions());
            using DbDataReader reader = workbook.OpenTable("Data", true, CancellationToken.None);
            Assert.Equal(3, reader.FieldCount);
            Assert.Equal("Id", reader.GetName(0));
            Assert.Equal("Value", reader.GetName(1));
            Assert.True(reader.Read());
            Assert.Equal(42, reader.GetInt32(0));
            Assert.Equal(7, reader.GetInt32(1));
            Assert.True(reader.IsDBNull(2));
            Assert.True(reader.Read());
            for (int column = 0; column < reader.FieldCount; column++) Assert.True(reader.IsDBNull(column));
            Assert.True(reader.Read());
            Assert.Equal(43, reader.GetInt32(0));
            Assert.True(reader.IsDBNull(1));
            Assert.Equal(9, reader.GetInt32(2));
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }
}
