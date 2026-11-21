using System.Data.Common;
using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Fact]
    public void XlsxMultiSheetReader_ReturnsEachSheetAsItsOwnReaderDoes_WhenALaterSheetNeedsTheSdkReader() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.XlsxNativeMultiSheet.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet first = document.AddSheet("First");
                first.Cell(1, 1).SetValue("Value");
                first.Cell(2, 1).SetValue(11);
                first.Cell(3, 1).SetValue(12);
                document.AddSheet("Second").Cell(1, 1).SetValue("Value");
                document.Save();
            }
            // Rows out of order are read by the SDK reader, not the native indexer.
            ReplaceZipEntry(path, "xl/worksheets/sheet2.xml", Encoding.UTF8.GetBytes("""
                <?xml version="1.0" encoding="UTF-8"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <sheetData>
                    <row r="1"><c r="A1" t="inlineStr"><is><t>Value</t></is></c></row>
                    <row r="3"><c r="A3"><v>23</v></c></row>
                    <row r="2"><c r="A2"><v>22</v></c></row>
                  </sheetData>
                </worksheet>
                """));
            using (XlsxTabularWorkbook native = XlsxTabularWorkbook.Open(path, new ExcelReaderOptions())) {
                Assert.Throws<XlsxTabularFastPathNotSupportedException>(() =>
                    native.OpenTable("Second", hasHeaderRow: true, CancellationToken.None).Dispose());
            }

            var expected = new List<List<object>>();
            for (int sheet = 0; sheet < 2; sheet++) {
                using ExcelWorkbookDataReader single = ExcelDocument.OpenReader(path, new ExcelReaderOptions { SheetIndex = sheet });
                expected.Add(ReadAllValues(single));
            }

            using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(path);
            Assert.Equal(new[] { "First", "Second" }, reader.SheetNames);
            Assert.Equal(expected[0], ReadAllValues(reader));
            Assert.True(reader.NextResult());
            Assert.Equal("Second", reader.CurrentSheetName);
            Assert.Equal(expected[1], ReadAllValues(reader));
            Assert.False(reader.NextResult());
        } finally {
            File.Delete(path);
        }

        static List<object> ReadAllValues(DbDataReader reader) {
            var values = new List<object>();
            for (int column = 0; column < reader.FieldCount; column++) values.Add(reader.GetName(column));
            while (reader.Read()) {
                for (int column = 0; column < reader.FieldCount; column++) values.Add(reader.GetValue(column));
            }
            return values;
        }
    }
}
