using System.Data;
using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData(8, false, false, false)]
    [InlineData(8, true, false, false)]
    [InlineData(65, false, false, false)]
    [InlineData(65, true, false, false)]
    [InlineData(8, false, true, false)]
    [InlineData(8, true, true, false)]
    [InlineData(65, false, true, false)]
    [InlineData(65, true, true, false)]
    [InlineData(8, false, false, true)]
    [InlineData(8, true, false, true)]
    [InlineData(65, false, false, true)]
    [InlineData(65, true, false, true)]
    [InlineData(8, false, true, true)]
    [InlineData(8, true, true, true)]
    [InlineData(65, false, true, true)]
    [InlineData(65, true, true, true)]
    public void Reader_DataTablePreservesSparseAndRepeatedRows(int width, bool utf16, bool inferTypes, bool startsLate) {
        string path = CreateCompactFastPathWorkbook();
        string lastColumn = width == 8 ? "I" : "BN";
        string encodingName = utf16 ? "utf-16" : "utf-8";
        const string rowThree = "<row r=\"3\"><c r=\"B3\"><v>30</v></c><c r=\"C3\" t=\"str\"><v>three</v></c><c r=\"D3\"><v>0</v></c></row>";
        const string rowSix = "<row r=\"6\"><c r=\"B6\"><v>60</v></c><c r=\"C6\" t=\"str\"><v>six</v></c></row>";
        string firstDataRows = startsLate ? rowSix + rowThree : rowThree + rowSix;
        string xml = $$"""
            <?xml version="1.0" encoding="{{encodingName}}"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
              <row r="2"><c r="B2" t="inlineStr"><is><t>Key</t></is></c></row>
              {{firstDataRows}}
              <row r="5"><c r="B5"><v>50</v></c><c r="{{lastColumn}}5"><v>500</v></c></row>
              <row r="3"><c r="B3"><v>31</v></c><c r="C3" t="inlineStr"><is/></c><c r="D3"><v>-0</v></c></row>
              <row r="8"><c r="B8"><v>80</v></c><c r="C8" t="str"><v>eight</v></c><c r="{{lastColumn}}8"><v>800</v></c></row>
              <row r="6"><c r="{{lastColumn}}6"><v>600</v></c></row>
              <row r="8"><c r="{{lastColumn}}8"/></row>
            </sheetData></worksheet>
            """;
        try {
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", (utf16 ? Encoding.Unicode : Encoding.UTF8).GetBytes(xml));
            using var owner = ExcelDocumentReader.Open(path, new ExcelReadOptions {
                NumericAsDecimal = false, InferDataTableColumnTypes = inferTypes
            });
            using var table = owner.GetSheet("Data").ReadRangeAsDataTable($"B2:{lastColumn}8", headersInFirstRow: true);

            Assert.Equal(width, table.Columns.Count);
            Assert.Equal("Key", table.Columns[0].ColumnName);
            Assert.Equal(inferTypes ? typeof(double) : typeof(object), table.Columns[0].DataType);
            Assert.Equal(6, table.Rows.Count);
            Assert.Equal(31d, table.Rows[0][0]);
            Assert.Equal(string.Empty, table.Rows[0][1]);
            Assert.Equal(0d, table.Rows[0][2]);
#if NET8_0_OR_GREATER
            // Object columns preserve the decoded value's bits. DataTable's typed
            // double storage normalizes signed zero, independently of this reader.
            if (!inferTypes) Assert.Equal(long.MinValue, BitConverter.DoubleToInt64Bits((double)table.Rows[0][2]));
#endif
            foreach (int row in new[] { 1, 4 }) {
                for (int column = 0; column < width; column++) Assert.True(table.Rows[row].IsNull(column));
            }
            Assert.Equal(50d, table.Rows[2][0]);
            Assert.Equal(500d, table.Rows[2][width - 1]);
            Assert.Equal(60d, table.Rows[3][0]);
            Assert.Equal("six", table.Rows[3][1]);
            Assert.Equal(600d, table.Rows[3][width - 1]);
            Assert.Equal(80d, table.Rows[5][0]);
            Assert.Equal("eight", table.Rows[5][1]);
            Assert.True(table.Rows[5].IsNull(width - 1));
            foreach (DataRow row in table.Rows) Assert.Equal(DataRowState.Added, row.RowState);
        } finally {
            File.Delete(path);
        }
    }
}
