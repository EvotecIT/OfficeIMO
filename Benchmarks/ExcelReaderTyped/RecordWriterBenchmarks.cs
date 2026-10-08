using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks;

/// <summary>Writes the upstream typed records using attributes and an explicit generated layout.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("WriteRecords")]
public class RecordWriterBenchmarks {
    private List<TypedRecord> _records = [];
    private List<MappedWriteRecord> _mapped = [];
    private readonly NativeWriterWorkload _native = new();
    private readonly ConfiguredXlsxWriterWorkload _xlsx = new();

    [ParamsSource(nameof(RowCounts))]
    public int RowCount { get; set; } = 50_000;
    [Params(ExcelFileFormat.Xlsx, ExcelFileFormat.Xlsb)]
    public ExcelFileFormat Format { get; set; }
    public IEnumerable<int> RowCounts() => new TypedReadBenchmarks().RowCounts();

    [GlobalSetup]
    public async Task SetupAsync() {
        BenchmarkInput.WriteDescription();
        if (Format == ExcelFileFormat.Xls && RowCount > 65_535) throw new ArgumentOutOfRangeException(nameof(RowCount));
        _records = Enumerable.Range(1, RowCount).Select(TypedWorkbookFixture.ExpectedRecord).ToList();
        _mapped = _records.Select(static value => new MappedWriteRecord {
            Name = value.Name, Id = value.Id, Date = value.Date, Value = value.Value,
        }).ToList();
        if (Format == ExcelFileFormat.Xlsx) await _xlsx.SetupAsync(RowCount, sharedStrings: false);
        else await _native.SetupAsync(RowCount, Format, recordWriterCapacity: true);
        using var output = new MemoryStream(4 * 1024 * 1024);
        await WriteRecordsAsync(output, mapped: false);
        Validate(output.ToArray());
        if (Format != ExcelFileFormat.Xls) {
            output.SetLength(0);
            await WriteRecordsAsync(output, mapped: true);
            Validate(output.ToArray());
        }
    }

    [Benchmark]
    public long OfficeIMOPublicTypedWrite() => Format == ExcelFileFormat.Xlsx
        ? _xlsx.OfficeIMO(sharedStrings: false, includeReferences: true) : _native.OfficeIMO();
    [Benchmark(Baseline = true)]
    public Task<long> ExcelReaderWriteRecords() => WriteAsync(mapped: false);

    internal Task<long> ExcelReaderMappedRecords() => WriteAsync(mapped: true);

    private async Task<long> WriteAsync(bool mapped) {
        await using var stream = new MemoryStream(4 * 1024 * 1024);
        await WriteRecordsAsync(stream, mapped);
        return stream.Length;
    }

    private async Task WriteRecordsAsync(Stream stream, bool mapped) {
        switch (Format) {
            case ExcelFileFormat.Xlsx:
                await using (var workbook = XlsxWorkbookWriter.Create(stream, leaveOpen: true)) {
                    await using var sheet = workbook.AddSheet("S1");
                    if (mapped) await sheet.WriteRecordsAsync(_mapped, ExcelRecordLayout.Generated<MappedWriteRecord>());
                    else await sheet.WriteRecordsAsync(_records, ExcelRecordLayout.FromAttributes<TypedRecord>());
                }
                break;
            case ExcelFileFormat.Xlsb:
                await using (var workbook = XlsbWorkbookWriter.Create(stream, leaveOpen: true)) {
                    await using var sheet = workbook.AddSheet("S1");
                    if (mapped) await sheet.WriteRecordsAsync(_mapped, ExcelRecordLayout.Generated<MappedWriteRecord>());
                    else await sheet.WriteRecordsAsync(_records, ExcelRecordLayout.FromAttributes<TypedRecord>());
                }
                break;
            case ExcelFileFormat.Xls:
                if (mapped) throw new InvalidOperationException("The upstream suite has no mapped XLS record method.");
                await using (var workbook = XlsWorkbookWriter.Create(stream, leaveOpen: true)) {
                    await using var sheet = workbook.AddSheet("S1");
                    await sheet.WriteRecordsAsync(_records, ExcelRecordLayout.FromAttributes<TypedRecord>());
                }
                break;
            default: throw new ArgumentOutOfRangeException(nameof(Format));
        }
    }

    private void Validate(byte[] bytes) {
        if (Format == ExcelFileFormat.Xlsx) WrittenWorkbookValidation.Validate(bytes, RowCount, officeIMO: false);
        else NativeWrittenWorkbookValidation.Validate(bytes, RowCount, Format, "ExcelReader typed records");
    }
}

/// <summary>Retains the typed XLS writer separately from qualified XLSX/XLSB comparisons.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("WriteRecordsXlsDiagnostic")]
public class XlsRecordWriterDiagnosticBenchmarks {
    private RecordWriterBenchmarks _workload = null!;
    [ParamsSource(nameof(RowCounts))]
    public int RowCount { get; set; } = 50_000;
    public IEnumerable<int> RowCounts() => new TypedReadBenchmarks().RowCounts();
    [GlobalSetup]
    public Task SetupAsync() {
        _workload = new RecordWriterBenchmarks { RowCount = RowCount, Format = ExcelFileFormat.Xls };
        return _workload.SetupAsync();
    }
    [Benchmark]
    public long OfficeIMOModelAndSave() => _workload.OfficeIMOPublicTypedWrite();
    [Benchmark]
    public Task<long> ExcelReaderTypedXlsWriterDiagnostic() => _workload.ExcelReaderWriteRecords();
}

/// <summary>The explicit layout used by the upstream mapped-record writer.</summary>
public sealed class MappedWriteRecord : IExcelRecordMap<MappedWriteRecord> {
    public string? Name { get; set; }
    public int Id { get; set; }
    public DateTime Date { get; set; }
    public double Value { get; set; }
    public static void ConfigureExcelRecordMap<TRow>(ExcelRecordMapBuilder<MappedWriteRecord, TRow> builder)
        where TRow : IRowWriter => builder.Column("Name", static (row, record) => row.Write(record.Name))
            .Column("Id", static (row, record) => row.Write(record.Id))
            .Column("Date", static (row, record) => row.Write(record.Date))
            .Column("Value", static (row, record) => row.Write(record.Value));
}

/// <summary>Separates the mapped XLSX/XLSB cases from the ordinary attribute layout.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("WriteMappedRecords")]
public class MappedRecordWriterBenchmarks {
    private RecordWriterBenchmarks _workload = null!;
    [ParamsSource(nameof(RowCounts))]
    public int RowCount { get; set; } = 50_000;
    [Params(ExcelFileFormat.Xlsx, ExcelFileFormat.Xlsb)]
    public ExcelFileFormat Format { get; set; }
    public IEnumerable<int> RowCounts() => new TypedReadBenchmarks().RowCounts();
    [GlobalSetup]
    public Task SetupAsync() {
        _workload = new RecordWriterBenchmarks { RowCount = RowCount, Format = Format };
        return _workload.SetupAsync();
    }
    [Benchmark]
    public long OfficeIMOPublicTypedWrite() => _workload.OfficeIMOPublicTypedWrite();
    [Benchmark(Baseline = true)]
    public Task<long> ExcelReaderMappedRecords() => _workload.ExcelReaderMappedRecords();
}
