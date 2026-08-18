using System.Data.Common;
using System.Text;
using System.Threading;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Fact]
    public void XlsxTabularWorkbook_ReadsCanonicalPackageWithoutSdkProjection() {
        string path = CreateCompactFastPathWorkbook();
        try {
            var options = new ExcelReadOptions();
            using var workbook = XlsxTabularWorkbook.Open(path, options);

            Assert.Equal(new[] { "Data" }, workbook.TableNames);
            using DbDataReader reader = workbook.OpenTable(
                "Data",
                hasHeaderRow: true,
                CancellationToken.None);
            Assert.True(reader.Read());
            Assert.Equal(42, reader.GetInt32(0));
            Assert.True(reader.Read());
            Assert.Equal(43, reader.GetInt32(0));
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_ReusesValidatedSharedStringIndexForCachedFormula() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.XlsxNativeSharedFormula.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Value");
                sheet.CellValue(2, 1, "Cached");
                document.Save();
            }

            const string entryName = "xl/worksheets/sheet1.xml";
            string worksheetXml = Encoding.UTF8.GetString(ReadZipEntry(path, entryName));
            int cellStart = worksheetXml.IndexOf("<c r=\"A2\"", StringComparison.Ordinal);
            Assert.True(cellStart >= 0);
            int cellContentStart = worksheetXml.IndexOf('>', cellStart) + 1;
            Assert.True(cellContentStart > cellStart);
            worksheetXml = worksheetXml.Insert(
                cellContentStart,
                "<f>CONCAT(\"Cache\",\"d\")</f>");
            ReplaceZipEntry(path, entryName, Encoding.UTF8.GetBytes(worksheetXml));

            using (DbDataReader cached = ExcelDocument.OpenDataReader(path)) {
                Assert.True(cached.Read());
                Assert.Equal("Cached", cached.GetString(0));
            }

            using (DbDataReader formula = ExcelDocument.OpenDataReader(
                       path,
                       new ExcelReadOptions { UseCachedFormulaResult = false })) {
                Assert.True(formula.Read());
                Assert.Equal("CONCAT(\"Cache\",\"d\")", formula.GetString(0));
            }
        } finally {
            File.Delete(path);
        }
    }
}
