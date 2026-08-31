#if NET8_0_OR_GREATER
using Apache.Arrow;
using Apache.Arrow.Arrays;
using OfficeIMO.Data.Arrow;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void DataReader_ArrowFastPathPreservesTypedValuesAndNulls() {
        DateTime expectedDate = new(2026, 8, 31, 14, 30, 0, DateTimeKind.Unspecified);
        using var memory = new MemoryStream();
        using (var document = ExcelDocument.Create(
                   memory,
                   new ExcelCreateOptions { PersistenceMode = OfficeIMO.DocumentPersistenceMode.SaveOnDispose })) {
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Name");
            sheet.CellValue(1, 2, "Id");
            sheet.CellValue(1, 3, "Amount");
            sheet.CellValue(1, 4, "When");
            sheet.CellValue(2, 1, "Alpha & Beta");
            sheet.CellValue(2, 2, 7);
            sheet.CellValue(2, 3, 12.5d);
            sheet.CellValue(2, 4, expectedDate);
            sheet.CellValue(3, 1, "Gamma");
            sheet.CellValue(3, 2, 8);
        }

        using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(memory.ToArray());
        RecordBatch batch = Assert.Single(reader.ReadArrowBatches(new ArrowReadOptions {
            BatchSize = 16,
            ColumnTypes = [typeof(string), typeof(long), typeof(double), typeof(DateTime)]
        }));
        using (batch) {
            Assert.Equal(2, batch.Length);
            Assert.Equal("Alpha & Beta", Assert.IsType<StringArray>(batch.Column(0)).GetString(0));
            Assert.Equal("Gamma", Assert.IsType<StringArray>(batch.Column(0)).GetString(1));
            Assert.Equal(7L, Assert.IsType<Int64Array>(batch.Column(1)).GetValue(0));
            Assert.Equal(8L, Assert.IsType<Int64Array>(batch.Column(1)).GetValue(1));
            Assert.Equal(12.5d, Assert.IsType<DoubleArray>(batch.Column(2)).GetValue(0));
            Assert.Null(Assert.IsType<DoubleArray>(batch.Column(2)).GetValue(1));
            Assert.Equal(expectedDate, Assert.IsType<TimestampArray>(batch.Column(3)).GetTimestamp(0)?.DateTime);
            Assert.Null(Assert.IsType<TimestampArray>(batch.Column(3)).GetTimestamp(1));
        }
    }
}
#endif
