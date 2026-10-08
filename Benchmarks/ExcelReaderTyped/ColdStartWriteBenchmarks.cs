using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Engines;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xlsx;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks;

/// <summary>Measures the first automatic object-to-column layout use for each writer.</summary>
[MemoryDiagnoser]
[SimpleJob(RunStrategy.ColdStart, launchCount: 16, warmupCount: 0, iterationCount: 1, invocationCount: 1)]
[BenchmarkCategory("ColdStartWrite")]
public class ColdStartWriteBenchmarks {
    private List<ColdStartRecord> _records = [];

    [GlobalSetup]
    public void Setup() {
        BenchmarkInput.WriteDescription();
        _records = ColdStartWorkload.Records();
    }

    [Benchmark(Baseline = true)]
    public async Task<long> ExcelReaderAutomaticLayout() {
        await using var stream = new MemoryStream(64 * 1024);
        await WritePeerAsync(stream);
        return stream.Length;
    }

    [Benchmark]
    public long OfficeIMOAutomaticLayout() {
        using var stream = new MemoryStream(64 * 1024);
        WriteOfficeIMO(stream);
        return stream.Length;
    }

    internal async Task ValidateAsync() {
        using var office = new MemoryStream(64 * 1024);
        WriteOfficeIMO(office);
        WrittenWorkbookValidation.Validate(office.ToArray(), ColdStartWorkload.Rows, officeIMO: true);
        using var peer = new MemoryStream(64 * 1024);
        await WritePeerAsync(peer);
        WrittenWorkbookValidation.Validate(peer.ToArray(), ColdStartWorkload.Rows, officeIMO: false);
    }

    private void WriteOfficeIMO(Stream stream) {
        using var document = ExcelDocument.Create();
        document.AddWorksheet("Data").InsertObjects(_records);
        document.Save(stream);
    }

    private async Task WritePeerAsync(Stream stream) {
        await using var workbook = XlsxWorkbookWriter.Create(stream, leaveOpen: true);
        await using var sheet = workbook.AddSheet("S1");
        await sheet.WriteRecordsAsync(_records, ExcelRecordLayout.FromAttributes<ColdStartRecord>());
    }
}
