using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Guards the single-pass object-row contract with a generated source large enough to make
/// accidental row buffering visible in both execution time and managed allocations.
/// </summary>
[MemoryDiagnoser]
public class ExcelGeneratedRowStreamingBenchmarks {
    private static readonly string[] Headers = ["Id", "Amount", "CreatedOn", "Active"];

    [Params(1_000_000)]
    public int RowCount { get; set; }

    [Benchmark]
    public int WriteRowsGenerated() {
        ExcelDataSetImportResult result = ExcelDocument.WriteRows(
            Stream.Null,
            GenerateRows(RowCount),
            Headers,
            static (writer, row) => writer
                .Write(row.Id)
                .Write(row.Amount)
                .Write(row.CreatedOn)
                .Write(row.Active));
        return result.RowCount;
    }

    private static IEnumerable<GeneratedRow> GenerateRows(int count) {
        var start = new DateTime(2026, 1, 1);
        for (int index = 0; index < count; index++) {
            yield return new GeneratedRow(
                index + 1,
                index * 1.25m,
                start.AddMinutes(index),
                (index & 1) == 0);
        }
    }

    private readonly record struct GeneratedRow(
        int Id,
        decimal Amount,
        DateTime CreatedOn,
        bool Active);
}
