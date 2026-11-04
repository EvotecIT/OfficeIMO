#if NET10_0_OR_GREATER
using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Measures complete XLSB exports through typed cell buffering. This diagnostic
/// compares OfficeIMO revisions; engine comparisons use the generated writer lanes.
/// </summary>
[MemoryDiagnoser]
public class ExcelBufferedXlsbWriteBenchmarks {
    [Params(50_000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup() => new ExcelReaderGeneratedWriteBenchmarks {
        Format = ExcelFileFormat.Xlsb, RowCount = RowCount
    }.Validate(Write());

    [Benchmark]
    public byte[] Write() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Data");
        for (int column = 0; column < ExcelReaderGeneratedBenchmarks.Headers.Length; column++) {
            sheet.Cell(1, column + 1).SetValue(ExcelReaderGeneratedBenchmarks.Headers[column]);
        }
        for (int index = 0; index < RowCount; index++) {
            var row = ExcelReaderGeneratedBenchmarks.CreateRecord(index);
            sheet.Cell(index + 2, 1).SetValue(row.Name);
            sheet.Cell(index + 2, 2).SetValue(row.Id);
            sheet.Cell(index + 2, 3).SetValue(row.CreatedOn);
            sheet.Cell(index + 2, 4).SetValue(row.Amount);
        }
        return document.ToBytes(ExcelFileFormat.Xlsb);
    }
}
#endif
