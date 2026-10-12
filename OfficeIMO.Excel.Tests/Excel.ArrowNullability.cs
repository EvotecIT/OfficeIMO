#if NET8_0_OR_GREATER
using Apache.Arrow;
using OfficeIMO.Data.Arrow;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void DataReader_RequiredArrowColumnRejectsNullFromTheWorkbookFastPath() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Id");
        sheet.CellValue(1, 2, "Name");
        sheet.CellValue(2, 2, "Alpha");
        sheet.CellValue(3, 1, 42);
        sheet.CellValue(3, 2, "Beta");
        byte[] workbook = document.ToBytes();

        using (ExcelWorkbookDataReader nullableReader = ExcelDocument.OpenDataReader(workbook)) {
            using RecordBatch batch = Assert.Single(nullableReader.ReadArrowBatches(new ArrowReadOptions {
                ColumnTypes = [typeof(long), typeof(string)]
            }));
            Assert.Null(Assert.IsType<Int64Array>(batch.Column(0)).GetValue(0));
            Assert.Equal(42L, Assert.IsType<Int64Array>(batch.Column(0)).GetValue(1));
        }

        using ExcelWorkbookDataReader requiredReader = ExcelDocument.OpenDataReader(workbook);
        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            requiredReader.ReadArrowBatches(new ArrowReadOptions {
                ColumnTypes = [typeof(long), typeof(string)],
                ColumnNullability = [false, true]
            }).Single());
        Assert.Contains("Required Arrow column 'Id'", exception.Message, StringComparison.Ordinal);
    }
}
#endif
