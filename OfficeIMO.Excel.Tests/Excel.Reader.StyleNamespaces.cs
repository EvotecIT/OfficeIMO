using System.Text;
using System.Threading;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData("number-format", false)]
    [InlineData("number-format", true)]
    [InlineData("cell-format", false)]
    [InlineData("cell-format", true)]
    [InlineData("nested-number-format", false)]
    [InlineData("nested-number-format", true)]
    [InlineData("nested-cell-format", false)]
    [InlineData("nested-cell-format", true)]
    public void Reader_UsesOnlyOwnedStyleDefinitions(string shadow, bool cancellable) {
        string path = CreateCompactFastPathWorkbook();
        using var cancellation = new CancellationTokenSource();
        try {
            string foreignNumberFormat = shadow == "number-format"
                ? "<p:payload><p:numFmt numFmtId=\"165\" formatCode=\"yyyy-mm-dd\"/></p:payload>"
                : shadow == "nested-number-format"
                    ? "<p:payload><numFmt numFmtId=\"165\" formatCode=\"yyyy-mm-dd\"/></p:payload>" : "";
            string foreignCellFormat = shadow == "cell-format" ? "<p:xf numFmtId=\"14\"/>"
                : shadow == "nested-cell-format" ? "<p:payload><xf numFmtId=\"14\"/></p:payload>" : "";
            string styles = "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
                + " xmlns:p=\"urn:private:style-payload\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"p\">"
                + "<numFmts count=\"1\"><numFmt numFmtId=\"165\" formatCode=\"0.00\"/></numFmts>"
                + "<fonts count=\"1\"><font><name val=\"Calibri\"/>" + foreignNumberFormat + "</font></fonts>"
                + "<cellXfs count=\"2\"><xf numFmtId=\"0\"/>" + foreignCellFormat
                + "<xf numFmtId=\"165\"/></cellXfs></styleSheet>";
            ReplaceZipEntry(path, "xl/styles.xml", Encoding.UTF8.GetBytes(styles));
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes("""
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
                  <row><c t="inlineStr"><is><t>Value</t></is></c></row>
                  <row><c s="1"><v>42</v></c></row>
                  <row><c s="1"><v>43</v></c></row>
                </sheetData></worksheet>
                """));
            using var reader = ExcelDocument.OpenReader(path, new ExcelReaderOptions {
                NumericAsDecimal = false,
                CancellationToken = cancellable ? cancellation.Token : CancellationToken.None
            });
            Assert.True(reader.Read());
            Assert.Equal(42d, Assert.IsType<double>(reader.GetValue(0)));
            Assert.True(reader.Read());
            Assert.Equal(43d, Assert.IsType<double>(reader.GetValue(0)));
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }
}
