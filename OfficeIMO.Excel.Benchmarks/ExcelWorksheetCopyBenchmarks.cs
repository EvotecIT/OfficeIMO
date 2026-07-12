using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;

namespace OfficeIMO.Excel.Benchmarks;

[MemoryDiagnoser]
[SimpleJob(RuntimeMoniker.Net80)]
public class ExcelWorksheetCopyBenchmarks {
    private byte[] _sourceWorkbookBytes = [];

    [Params(100, 2500, 25000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        var rows = ExcelBenchmarkScenarioFactory.CreateSalesRecords(RowCount);
        _sourceWorkbookBytes = ExcelBenchmarkScenarioFactory.CreateWorkbookBytes(rows);
    }

    [Benchmark(Baseline = true)]
    public int PackageCopy()
        => CopyWorksheet(ExcelWorksheetCopyMode.Package);

    [Benchmark]
    public int ValuesCopy()
        => CopyWorksheet(ExcelWorksheetCopyMode.Values);

    private int CopyWorksheet(ExcelWorksheetCopyMode copyMode) {
        using var sourceStream = new MemoryStream(_sourceWorkbookBytes, writable: false);
        using var sourceDocument = ExcelDocument.Load(sourceStream, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
        using var targetStream = new MemoryStream();
        using (var targetDocument = ExcelDocument.Create(targetStream)) {
            targetDocument.CopyWorksheetFrom(
                sourceDocument,
                "Data",
                "DataCopy",
                SheetNameValidationMode.Sanitize,
                new ExcelWorksheetCopyOptions { CopyMode = copyMode });
            targetDocument.Save(targetStream);
        }

        return checked((int)targetStream.Length);
    }
}
