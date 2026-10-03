using System.Threading;
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.Excel.Tests;

/// <summary>Runs the blocking factory-overlap probe apart from other test collections.</summary>
[CollectionDefinition(Name, DisableParallelization = true)]
public sealed class ExcelParallelProjectionConcurrencyCollection {
    public const string Name = "Excel parallel projection concurrency";
}

[Collection(ExcelParallelProjectionConcurrencyCollection.Name)]
public sealed class ExcelParallelProjectionConcurrencyTests {
    [Fact]
    public void RowsAsParallel_FactoryActuallyRunsConcurrentlyAndPreservesOrder() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "OrderId");
        for (int index = 0; index < 8; index++) {
            sheet.CellValue(index + 2, 1, index);
        }

        using var firstWorkers = new Barrier(2);
        int calls = 0;
        int[] rows = sheet.RowsAsParallel(
            "A1:A9",
            factory: record => {
                if (Interlocked.Increment(ref calls) <= 2) {
                    Assert.True(firstWorkers.SignalAndWait(TimeSpan.FromSeconds(10)));
                }
                return record.GetInt32(0);
            },
            parallelOptions: new ParallelRowMappingOptions {
                MaxDegreeOfParallelism = 2,
                BatchSize = 1
            }).ToArray();

        Assert.Equal(Enumerable.Range(0, 8), rows);
    }
}
