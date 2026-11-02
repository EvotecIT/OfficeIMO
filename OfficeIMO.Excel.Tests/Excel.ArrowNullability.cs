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
        ExcelSheet sheet = document.AddSheet("Data");
        sheet.Cell(1, 1).SetValue("Id");
        sheet.Cell(1, 2).SetValue("Name");
        sheet.Cell(2, 2).SetValue("Alpha");
        sheet.Cell(3, 1).SetValue(42);
        sheet.Cell(3, 2).SetValue("Beta");
        byte[] workbook = document.ToBytes();

        using (ExcelWorkbookDataReader nullableReader = ExcelDocument.OpenReader(workbook)) {
            using RecordBatch batch = Assert.Single(nullableReader.ReadArrowBatches(new ArrowReadOptions {
                ColumnTypes = [typeof(long), typeof(string)]
            }));
            Assert.Null(Assert.IsType<Int64Array>(batch.Column(0)).GetValue(0));
            Assert.Equal(42L, Assert.IsType<Int64Array>(batch.Column(0)).GetValue(1));
        }

        using ExcelWorkbookDataReader requiredReader = ExcelDocument.OpenReader(workbook);
        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            requiredReader.ReadArrowBatches(new ArrowReadOptions {
                ColumnTypes = [typeof(long), typeof(string)],
                ColumnNullability = [false, true]
            }).Single());
        Assert.Contains("Required Arrow column 'Id'", exception.Message, StringComparison.Ordinal);
    }
}
#endif
