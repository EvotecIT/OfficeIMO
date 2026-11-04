#if NET10_0_OR_GREATER
using BenchmarkDotNet.Attributes;
using SpreadCheetah;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>Three native XLSX writers producing complete packages with the same four typed columns.</summary>
[MemoryDiagnoser]
public class ExcelReaderGeneratedXlsxWriteBenchmarks {
    private readonly ExcelReaderGeneratedWriteBenchmarks _workload = new() { Format = ExcelFileFormat.Xlsx, RowCount = 50_000 };

    [GlobalSetup]
    public void Setup() {
        _workload.SetupOfficeIMO();
        _workload.SetupExcelReader();
        _workload.Validate(SpreadCheetah());
    }

    [Benchmark(Baseline = true)] public byte[] OfficeIMO() => _workload.OfficeIMO();
    [Benchmark] public byte[] ExcelReaderNet() => _workload.ExcelReaderNet();
    [Benchmark] public byte[] SpreadCheetah() => WriteSpreadCheetah().GetAwaiter().GetResult();

    private async Task<byte[]> WriteSpreadCheetah() {
        using var output = new MemoryStream();
        await using var writer = await global::SpreadCheetah.Spreadsheet.CreateNewAsync(output);
        await writer.StartWorksheetAsync("Data");
        var cells = new DataCell[4];
        for (int column = 0; column < cells.Length; column++) cells[column] = new DataCell(ExcelReaderGeneratedBenchmarks.Headers[column]);
        await writer.AddRowAsync(cells);
        for (int index = 0; index < _workload.RowCount; index++) {
            var value = ExcelReaderGeneratedBenchmarks.CreateRecord(index);
            cells[0] = new DataCell(value.Name);
            cells[1] = new DataCell(value.Id);
            cells[2] = new DataCell(value.CreatedOn);
            cells[3] = new DataCell(value.Amount);
            await writer.AddRowAsync(cells);
        }
        await writer.FinishAsync();
        return output.ToArray();
    }
}
#endif
