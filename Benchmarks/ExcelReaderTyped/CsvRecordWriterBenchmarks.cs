using System.Globalization;
using System.Text;
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Csv;
using OfficeIMO.CSV;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks;

/// <summary>Uses the original automatic and generated CSV record layouts with complete output proof.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("CsvWriteRecords")]
public class CsvRecordWriterBenchmarks {
    private List<TypedRecord> _records = [];
    private List<MappedWriteRecord> _mapped = [];
    [ParamsSource(nameof(RowCounts))]
    public int RowCount { get; set; } = 50_000;
    [Params(false, true)]
    public bool Mapped { get; set; }
    public IEnumerable<int> RowCounts() => new TypedReadBenchmarks().RowCounts();
    [GlobalSetup]
    public async Task SetupAsync() {
        BenchmarkInput.WriteDescription();
        _records = Enumerable.Range(1, RowCount).Select(TypedWorkbookFixture.ExpectedRecord).ToList();
        _mapped = _records.Select(static record => new MappedWriteRecord {
            Name = record.Name, Id = record.Id, Date = record.Date, Value = record.Value,
        }).ToList();
        using var peer = new MemoryStream(4 * 1024 * 1024);
        await WritePeerAsync(peer);
        Validate(peer, "ExcelReader");
        using var office = new MemoryStream(4 * 1024 * 1024);
        WriteOffice(office);
        Validate(office, "OfficeIMO");
    }
    [Benchmark(Baseline = true)]
    public async Task<long> ExcelReaderRecordLayout() {
        await using var stream = new MemoryStream(4 * 1024 * 1024);
        await WritePeerAsync(stream);
        return stream.Length;
    }
    [Benchmark]
    public long OfficeIMOWriteObjects() {
        using var stream = new MemoryStream(4 * 1024 * 1024);
        WriteOffice(stream);
        return stream.Length;
    }
    private async Task WritePeerAsync(Stream stream) {
        await using var workbook = CsvWorkbookWriter.Create(stream, leaveOpen: true);
        await using var sheet = workbook.AddSheet("S1");
        if (Mapped) await sheet.WriteRecordsAsync(_mapped, ExcelRecordLayout.Generated<MappedWriteRecord>());
        else await sheet.WriteRecordsAsync(_records, ExcelRecordLayout.FromAttributes<TypedRecord>());
    }
    private void WriteOffice(Stream stream) {
        using var text = new StreamWriter(stream, new UTF8Encoding(false), 1024, leaveOpen: true);
        var options = new CsvSaveOptions { Culture = CultureInfo.InvariantCulture, DateTimeFormat = "O" };
        if (Mapped) CsvDocument.WriteObjects(text, _mapped, options);
        else CsvDocument.WriteObjects(text, _records, options);
    }
    private void Validate(MemoryStream stream, string engine) {
        stream.Position = 0;
        using (var reader = CsvDocument.OpenDataReader(stream)) CsvBenchmarkFixture.Validate(reader, RowCount, headers: true);
        stream.Position = 0;
        using var text = new StreamReader(stream, Encoding.UTF8, true, 1024, leaveOpen: true);
        using var independent = global::Sylvan.Data.Csv.CsvDataReader.Create(text,
            new global::Sylvan.Data.Csv.CsvDataReaderOptions { Culture = CultureInfo.InvariantCulture });
        CsvBenchmarkFixture.Validate(independent, RowCount, headers: true);
        Console.WriteLine($"Qualified CSV record writer={engine}; mapped={Mapped}; rows={RowCount}; bytes={stream.Length}; "
            + "every header/name/ID/date/value/order verified by OfficeIMO and Sylvan. OfficeIMO uses ordinary object mapping.");
    }
}
