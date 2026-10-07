using System.Data;
using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Reader_DataTableConverterKeepsLateRepeatedRows(bool inferTypes, bool utf16) {
        string path = CreateCompactFastPathWorkbook();
        string encodingName = utf16 ? "utf-16" : "utf-8";
        string xml = $$"""
            <?xml version="1.0" encoding="{{encodingName}}"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
              <row r="1"><c r="A1" t="str"><v>old</v></c></row>
              <row r="2"><c r="A2" t="str"><v>two</v></c></row>
              <row r="3"><c r="A3" t="str"><v>outside</v></c></row>
              <row r="4"/>
              <row r="1"><c r="A1" t="str"><v>new</v></c></row>
            </sheetData></worksheet>
            """;
        try {
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", (utf16 ? Encoding.Unicode : Encoding.UTF8).GetBytes(xml));
            var options = new ExcelReadOptions {
                InferDataTableColumnTypes = inferTypes,
                CellValueConverter = static _ => ExcelCellValue.NotHandled
            };
            using var owner = ExcelDocumentReader.Open(path, options);
            using DataTable table = owner.GetSheet("Data").ReadRangeAsDataTable("A1:A2", headersInFirstRow: false);

            Assert.Equal(2, table.Rows.Count);
            Assert.Equal("new", table.Rows[0][0]);
            Assert.Equal("two", table.Rows[1][0]);
        } finally {
            File.Delete(path);
        }
    }
}
