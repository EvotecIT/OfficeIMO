using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Writer.Xlsx;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks;

/// <summary>Writes the same records with shared strings selected in both libraries.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("WriteSharedStrings")]
public class SharedStringWriterBenchmarks {
    private readonly ConfiguredXlsxWriterWorkload _workload = new();
    [ParamsSource(nameof(RowCounts))]
    public int RowCount { get; set; } = 50_000;
    public IEnumerable<int> RowCounts() => new TypedReadBenchmarks().RowCounts();
    [GlobalSetup]
    public Task SetupAsync() => _workload.SetupAsync(RowCount, sharedStrings: true);
    [Benchmark]
    public long OfficeIMOSharedStrings() => _workload.OfficeIMO(sharedStrings: true, includeReferences: true);
    [Benchmark]
    public long OfficeIMOSharedStringsWithoutReferences() => _workload.OfficeIMO(sharedStrings: true, includeReferences: false);
    [Benchmark(Baseline = true)]
    public Task<long> ExcelReaderWriterSharedStrings() => _workload.ExcelReader(sharedStrings: true);
}

/// <summary>Separates the public coordinate-omission option from the ordinary writer.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("WriteWithoutReferences")]
public class CompactWriterBenchmarks {
    private readonly ConfiguredXlsxWriterWorkload _workload = new();
    [ParamsSource(nameof(RowCounts))]
    public int RowCount { get; set; } = 50_000;
    public IEnumerable<int> RowCounts() => new TypedReadBenchmarks().RowCounts();
    [GlobalSetup]
    public Task SetupAsync() => _workload.SetupAsync(RowCount, sharedStrings: false);
    [Benchmark]
    public long OfficeIMOWithoutReferences() => _workload.OfficeIMO(sharedStrings: false, includeReferences: false);
    [Benchmark(Baseline = true)]
    public Task<long> ExcelReaderWriter() => _workload.ExcelReader(sharedStrings: false);
    [Benchmark]
    public Task<long> ExcelReaderWriterPrefetch() => _workload.ExcelReader(sharedStrings: false, prefetch: true);
}

internal sealed class ConfiguredXlsxWriterWorkload {
    private List<TypedRecord> _records = [];

    internal async Task SetupAsync(int rowCount, bool sharedStrings) {
        BenchmarkInput.WriteDescription();
        _records = Enumerable.Range(1, rowCount).Select(TypedWorkbookFixture.ExpectedRecord).ToList();
        foreach (bool references in new[] { true, false }) {
            using var stream = new MemoryStream(4 * 1024 * 1024);
            WriteOfficeIMO(stream, sharedStrings, references);
            WrittenWorkbookValidation.Validate(stream.ToArray(), rowCount, officeIMO: true,
                sharedStrings: sharedStrings, includeReferences: references);
        }
        foreach (bool prefetch in sharedStrings ? new[] { false } : new[] { false, true }) {
            using var stream = new MemoryStream(4 * 1024 * 1024);
            await WriteExcelReaderAsync(stream, sharedStrings, prefetch);
            WrittenWorkbookValidation.Validate(stream.ToArray(), rowCount, officeIMO: false, sharedStrings: sharedStrings);
        }
    }

    internal long OfficeIMO(bool sharedStrings, bool includeReferences) {
        using var stream = new MemoryStream(4 * 1024 * 1024);
        WriteOfficeIMO(stream, sharedStrings, includeReferences);
        return stream.Length;
    }

    internal async Task<long> ExcelReader(bool sharedStrings, bool prefetch = false) {
        await using var stream = new MemoryStream(4 * 1024 * 1024);
        await WriteExcelReaderAsync(stream, sharedStrings, prefetch);
        return stream.Length;
    }

    private void WriteOfficeIMO(Stream stream, bool sharedStrings, bool includeReferences) =>
        ExcelDocument.WriteRows(stream, _records, ["Name", "Id", "Date", "Value"], static (row, record) => {
            row.Write(record.Name);
            row.Write(record.Id);
            row.Write(record.Date);
            row.Write(record.Value);
        }, new ExcelTabularWriteOptions { UseSharedStrings = sharedStrings, IncludeCellReferences = includeReferences });

    private async Task WriteExcelReaderAsync(Stream stream, bool sharedStrings, bool prefetch) {
        await using var workbook = XlsxWorkbookWriter.Create(stream, leaveOpen: true,
            options: sharedStrings || prefetch ? new XlsxWriterOptions { UseSharedStrings = sharedStrings, PrefetchWrite = prefetch } : null);
        XlsxSheetWriter sheet = workbook.AddSheet("S1");
        using (XlsxRowWriter header = sheet.StartRow()) {
            header.Write("Name"); header.Write("Id"); header.Write("Date"); header.Write("Value");
        }
        for (int index = 0; index < _records.Count; index++) {
            TypedRecord record = _records[index];
            using XlsxRowWriter row = sheet.StartRow();
            row.Write(record.Name); row.Write(record.Id); row.Write(record.Date); row.Write(record.Value);
        }
        await sheet.EndAsync();
        await workbook.EndAsync();
    }
}
