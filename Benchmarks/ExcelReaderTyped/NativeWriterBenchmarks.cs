using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Core.Writer.Xlsb;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks;

/// <summary>Includes ordinary model construction and native binary save in the measured operation.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("WriteNativeBinary")]
public class NativeWriterBenchmarks {
    private readonly NativeWriterWorkload _workload = new();
    [ParamsSource(nameof(RowCounts))]
    public int RowCount { get; set; } = 50_000;
    [Params(ExcelFileFormat.Xlsb)]
    public ExcelFileFormat Format { get; set; }
    public IEnumerable<int> RowCounts() => new TypedReadBenchmarks().RowCounts();
    [GlobalSetup]
    public Task SetupAsync() => _workload.SetupAsync(RowCount, Format);
    [Benchmark]
    public long OfficeIMOModelAndSave() => _workload.OfficeIMO();
    [Benchmark(Baseline = true)]
    public Task<long> ExcelReaderWriter() => _workload.ExcelReader();
}

/// <summary>Retains original XLS costs without claiming equivalent valid Excel output.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("WriteXlsOriginalDiagnostic")]
public class NativeXlsWriterDiagnosticBenchmarks {
    private readonly NativeWriterWorkload _workload = new();
    [ParamsSource(nameof(RowCounts))]
    public int RowCount { get; set; } = 50_000;
    public IEnumerable<int> RowCounts() => new TypedReadBenchmarks().RowCounts();
    [GlobalSetup]
    public Task SetupAsync() => _workload.SetupAsync(RowCount, ExcelFileFormat.Xls);
    [Benchmark]
    public long OfficeIMOModelAndSave() => _workload.OfficeIMO();
    [Benchmark]
    public Task<long> ExcelReaderXlsWriterDiagnostic() => _workload.ExcelReader();
}

/// <summary>Retains the upstream XLSB shared-string and write-prefetch operations.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("WriteXlsbOptions")]
public class ConfiguredXlsbWriterBenchmarks {
    private readonly NativeWriterWorkload _workload = new();
    [ParamsSource(nameof(RowCounts))]
    public int RowCount { get; set; } = 50_000;
    public IEnumerable<int> RowCounts() => new TypedReadBenchmarks().RowCounts();
    [GlobalSetup]
    public async Task SetupAsync() {
        await _workload.SetupAsync(RowCount, ExcelFileFormat.Xlsb);
        await _workload.ValidateXlsbOptionAsync(sharedStrings: true, prefetch: false);
        await _workload.ValidateXlsbOptionAsync(sharedStrings: false, prefetch: true);
#if OFFICEIMO_BENCHMARK_NEW_APIS
        _workload.ValidateOfficeXlsbOption(sharedStrings: false);
        _workload.ValidateOfficeXlsbOption(sharedStrings: true);
#endif
    }
    [Benchmark]
    public long OfficeIMOModelAndSave() => _workload.OfficeIMO();
#if OFFICEIMO_BENCHMARK_NEW_APIS
    [Benchmark]
    public long OfficeIMOModelAndSaveInlineStrings() => _workload.OfficeIMOConfiguredXlsb(sharedStrings: false);
    [Benchmark]
    public long OfficeIMOModelAndSaveSharedStrings() => _workload.OfficeIMOConfiguredXlsb(sharedStrings: true);
#endif
    [Benchmark(Baseline = true)]
    public Task<long> ExcelReaderXlsbWriterSharedStrings() => _workload.ExcelReader(sharedStrings: true);
    [Benchmark]
    public Task<long> ExcelReaderXlsbWriterPrefetch() => _workload.ExcelReader(prefetch: true);
}

internal sealed class NativeWriterWorkload {
    private List<TypedRecord> _records = [];
    private ExcelFileFormat _format;
    private int _capacity;
    private int Capacity => _capacity;

    internal async Task SetupAsync(int rowCount, ExcelFileFormat format, bool recordWriterCapacity = false) {
        BenchmarkInput.WriteDescription();
        if (format == ExcelFileFormat.Xls && rowCount > 65_535)
            throw new ArgumentOutOfRangeException(nameof(rowCount), "The XLS header and data must fit one BIFF8 worksheet.");
        _format = format;
        _capacity = (format == ExcelFileFormat.Xls && !recordWriterCapacity ? 16 : 4) * 1024 * 1024;
        _records = Enumerable.Range(1, rowCount).Select(TypedWorkbookFixture.ExpectedRecord).ToList();
        using var office = new MemoryStream(Capacity);
        WriteOfficeIMO(office);
        NativeWrittenWorkbookValidation.Validate(office.ToArray(), rowCount, format, "OfficeIMO");
        using var peer = new MemoryStream(Capacity);
        await WriteExcelReaderAsync(peer, false, false);
        NativeWrittenWorkbookValidation.Validate(peer.ToArray(), rowCount, format, "ExcelReader");
    }

    internal async Task ValidateXlsbOptionAsync(bool sharedStrings, bool prefetch) {
        using var stream = new MemoryStream(Capacity);
        await WriteExcelReaderAsync(stream, sharedStrings, prefetch);
        NativeWrittenWorkbookValidation.Validate(stream.ToArray(), _records.Count, _format,
            sharedStrings ? "ExcelReader shared strings" : "ExcelReader write prefetch", sharedStrings);
    }

    internal long OfficeIMO() {
        using var stream = new MemoryStream(Capacity);
        WriteOfficeIMO(stream);
        return stream.Length;
    }

    private void WriteOfficeIMO(Stream stream, ExcelSaveOptions? options = null) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.InsertObjects(_records, ("Name", static record => record.Name), ("Id", static record => record.Id),
            ("Date", static record => record.Date), ("Value", static record => record.Value));
        document.Save(stream, _format, options);
    }

#if OFFICEIMO_BENCHMARK_NEW_APIS
    internal void ValidateOfficeXlsbOption(bool sharedStrings) {
        if (_format != ExcelFileFormat.Xlsb) throw new InvalidOperationException("Configured storage is XLSB only.");
        using var stream = new MemoryStream(Capacity);
        WriteOfficeIMO(stream, new ExcelSaveOptions { XlsbUseSharedStrings = sharedStrings });
        NativeWrittenWorkbookValidation.Validate(stream.ToArray(), _records.Count, _format,
            sharedStrings ? "OfficeIMO shared strings" : "OfficeIMO inline strings", sharedStrings);
    }
    internal long OfficeIMOConfiguredXlsb(bool sharedStrings) {
        using var stream = new MemoryStream(Capacity);
        WriteOfficeIMO(stream, new ExcelSaveOptions { XlsbUseSharedStrings = sharedStrings });
        return stream.Length;
    }
#endif

    internal async Task<long> ExcelReader(bool sharedStrings = false, bool prefetch = false) {
        await using var stream = new MemoryStream(Capacity);
        await WriteExcelReaderAsync(stream, sharedStrings, prefetch);
        return stream.Length;
    }

    private Task WriteExcelReaderAsync(Stream stream, bool sharedStrings, bool prefetch) =>
        _format == ExcelFileFormat.Xls ? WriteXlsAsync(stream) : WriteXlsbAsync(stream, sharedStrings, prefetch);

    private async Task WriteXlsAsync(Stream stream) {
        await using var workbook = XlsWorkbookWriter.Create(stream, leaveOpen: true);
        XlsSheetWriter sheet = workbook.AddSheet("S1");
        using (XlsRowWriter header = sheet.StartRow()) {
            header.Write("Name"); header.Write("Id"); header.Write("Date"); header.Write("Value");
        }
        for (int index = 0; index < _records.Count; index++) {
            TypedRecord record = _records[index];
            using XlsRowWriter row = sheet.StartRow();
            row.Write(record.Name); row.Write(record.Id); row.Write(record.Date); row.Write(record.Value);
        }
        sheet.End();
        await workbook.EndAsync();
    }

    private async Task WriteXlsbAsync(Stream stream, bool sharedStrings, bool prefetch) {
        await using var workbook = XlsbWorkbookWriter.Create(stream, leaveOpen: true,
            options: sharedStrings || prefetch ? new XlsbWriterOptions { UseSharedStrings = sharedStrings, PrefetchWrite = prefetch } : null);
        XlsbSheetWriter sheet = workbook.AddSheet("S1");
        ReadOnlySpan<XlsbCell> header = [XlsbCell.Create("Name"), XlsbCell.Create("Id"), XlsbCell.Create("Date"), XlsbCell.Create("Value")];
        sheet.WriteRow(header);
        XlsbCell[] row = new XlsbCell[4];
        for (int index = 0; index < _records.Count; index++) {
            TypedRecord record = _records[index];
            row[0] = XlsbCell.Create(record.Name); row[1] = XlsbCell.Create(record.Id);
            row[2] = XlsbCell.Create(record.Date); row[3] = XlsbCell.Create(record.Value);
            sheet.WriteRow(row);
        }
        await sheet.EndAsync();
        await workbook.EndAsync();
    }
}
