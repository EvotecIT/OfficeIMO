using System.Text;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Reader_PublicUsedRangePreservesLateFragmentsWithExplicitAndImplicitRows(bool implicitRows, bool lateUpdates) {
        string path = CreateCompactFastPathWorkbook();
        try {
            string Row(int index, string cells) => implicitRows
                ? $"<row>{cells}</row>" : $"<row r=\"{index}\">{cells}</row>";
            string suffix = lateUpdates
                ? Row(2, "<c r=\"B2\"/>") + Row(3, "<c r=\"B3\" t=\"str\"><v>updated</v></c>")
                : string.Empty;
            string xml = "<?xml version=\"1.0\" encoding=\"utf-16\"?>"
                + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
                + Row(1, "<c r=\"A1\" t=\"str\"><v>Id</v></c><c r=\"B1\" t=\"str\"><v>Note</v></c>")
                + Row(2, "<c r=\"A2\"><v>7</v></c><c r=\"B2\" t=\"str\"><v>initial</v></c>")
                + Row(3, "<c r=\"A3\"><v>8</v></c><c r=\"B3\" t=\"str\"><v>original</v></c>")
                + Row(5000, "<c r=\"A5000\"><v>9</v></c>") + suffix
                + "</sheetData></worksheet>";
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.Unicode.GetBytes(xml));
            using var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions { InferSchema = false });
            Assert.Equal(2, reader.FieldCount);
            Assert.Equal("Id", reader.GetName(0));
            Assert.Equal("Note", reader.GetName(1));
            int count = 0;
            while (reader.Read()) {
                count++;
                if (count == 1) {
                    Assert.Equal(7, reader.GetInt32(0));
                    if (lateUpdates) Assert.True(reader.IsDBNull(1));
                    else Assert.Equal("initial", reader.GetString(1));
                } else if (count == 2) {
                    Assert.Equal(8, reader.GetInt32(0));
                    Assert.Equal(lateUpdates ? "updated" : "original", reader.GetString(1));
                } else if (count == 4999) {
                    Assert.Equal(9, reader.GetInt32(0));
                    Assert.True(reader.IsDBNull(1));
                } else {
                    Assert.True(reader.IsDBNull(0));
                    Assert.True(reader.IsDBNull(1));
                }
            }
            Assert.Equal(4999, count);
            Assert.False(reader.NextResult());
        } finally {
            File.Delete(path);
        }
    }
}
