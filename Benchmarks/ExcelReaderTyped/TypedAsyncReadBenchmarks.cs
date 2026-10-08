using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Parser;
using OfficeIMO.Data;
using Sylvan.Data;
using Sylvan.Data.Excel;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks;

/// <summary>Compares the existing public async enumeration APIs on identical typed XLSX inputs.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("TypedAsync")]
public class TypedAsyncReadBenchmarks {
    private readonly TypedReadBenchmarks _workload = new();

    [ParamsSource(nameof(RowCounts))]
    public int RowCount { get; set; } = 50_000;

    [ParamsSource(nameof(Shapes))]
    public string Shape { get; set; } = "Original";

    public IEnumerable<int> RowCounts() => _workload.RowCounts();
    public IEnumerable<string> Shapes() => _workload.Shapes();

    [GlobalSetup]
    public async Task SetupAsync() {
        _workload.RowCount = RowCount;
        _workload.Shape = Shape;
        await _workload.SetupAsync();
        await ValidateAsync(ReadOfficeIMOForValidation());
        await ValidateAsync(ReadExcelReaderForValidation());
        await ValidateAsync(ReadSylvanForValidation());
        await OfficeIMOTypedAsync();
        await ExcelReaderTypedAsync();
        await SylvanTypedAsync();
        Console.WriteLine($"Validated async typed {Shape}: rows={RowCount}; all fields checked.");
    }

    [Benchmark]
    public async Task<long> OfficeIMOTypedAsync() {
        // Opening memory input is synchronous; use the reader's existing public async projection.
        await using var reader = ExcelDocument.OpenDataReader(_workload.Workbook, new ExcelReadOptions { HasHeaderRow = true });
        long sum = 0;
        int count = 0;
        await foreach (TypedRecord record in reader.RowsAsAsync<TypedRecord>()) {
            sum = unchecked(sum + TypedWorkbookFixture.Accumulate(record));
            count++;
        }
        return _workload.Check(sum, count);
    }

    [Benchmark(Baseline = true)]
    public async Task<long> ExcelReaderTypedAsync() {
        await using var stream = new MemoryStream(_workload.Workbook, writable: false);
        await using var workbook = await ExcelReaderApi.FromXlsxAsync(stream);
        long sum = 0;
        int count = 0;
        await foreach (TypedRecord record in ExcelParser.FromAttributes<TypedRecord>().Parse(workbook.FirstSheet)) {
            sum = unchecked(sum + TypedWorkbookFixture.Accumulate(record));
            count++;
        }
        return _workload.Check(sum, count);
    }

    [Benchmark]
    public async Task<long> SylvanTypedAsync() {
        await using var stream = new MemoryStream(_workload.Workbook, writable: false);
        await using var reader = await global::Sylvan.Data.Excel.ExcelDataReader.CreateAsync(stream,
            ExcelWorkbookType.ExcelXml, new ExcelDataReaderOptions());
        long sum = 0;
        int count = 0;
        await foreach (TypedRecord record in reader.GetRecordsAsync<TypedRecord>()) {
            sum = unchecked(sum + TypedWorkbookFixture.Accumulate(record));
            count++;
        }
        return _workload.Check(sum, count);
    }

    private async Task ValidateAsync(IAsyncEnumerable<TypedRecord> records) {
        int count = 0;
        await foreach (TypedRecord record in records) TypedWorkbookFixture.ValidateRecord(record, ++count);
        if (count != RowCount) throw new InvalidDataException("Incorrect async typed row count.");
    }

    private async IAsyncEnumerable<TypedRecord> ReadOfficeIMOForValidation() {
        await using var reader = ExcelDocument.OpenDataReader(_workload.Workbook, new ExcelReadOptions { HasHeaderRow = true });
        await foreach (TypedRecord record in reader.RowsAsAsync<TypedRecord>()) yield return record;
    }

    private async IAsyncEnumerable<TypedRecord> ReadExcelReaderForValidation() {
        await using var stream = new MemoryStream(_workload.Workbook, writable: false);
        await using var workbook = await ExcelReaderApi.FromXlsxAsync(stream);
        await foreach (TypedRecord record in ExcelParser.FromAttributes<TypedRecord>().Parse(workbook.FirstSheet)) yield return record;
    }

    private async IAsyncEnumerable<TypedRecord> ReadSylvanForValidation() {
        await using var stream = new MemoryStream(_workload.Workbook, writable: false);
        await using var reader = await global::Sylvan.Data.Excel.ExcelDataReader.CreateAsync(stream,
            ExcelWorkbookType.ExcelXml, new ExcelDataReaderOptions());
        await foreach (TypedRecord record in reader.GetRecordsAsync<TypedRecord>()) yield return record;
    }
}
