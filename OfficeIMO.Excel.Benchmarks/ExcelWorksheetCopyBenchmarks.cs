using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Excel.Benchmarks;

[MemoryDiagnoser]
public class ExcelWorksheetCopyBenchmarks {
    private byte[] _sourceWorkbookBytes = [];

    [Params(100, 2500, 25000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        var rows = ExcelBenchmarkScenarioFactory.CreateSalesRecords(RowCount);
        _sourceWorkbookBytes = ExcelBenchmarkScenarioFactory.CreateWorkbookBytes(rows);
        ExcelSalesOutputValidator.ValidateWorkbook(PackageCopy(), rows, sheetName: "DataCopy");
        ExcelSalesOutputValidator.ValidateWorkbook(ValuesCopy(), rows, sheetName: "DataCopy");
    }

    [Benchmark(Baseline = true)]
    public byte[] PackageCopy()
        => CopyWorksheet(ExcelWorksheetCopyMode.Package);

    [Benchmark]
    public byte[] ValuesCopy()
        => CopyWorksheet(ExcelWorksheetCopyMode.Values);

    private byte[] CopyWorksheet(ExcelWorksheetCopyMode copyMode) {
        using var sourceStream = new MemoryStream(_sourceWorkbookBytes, writable: false);
        using var sourceDocument = ExcelDocument.Load(sourceStream, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
        using var targetStream = new MemoryStream();
        using (var targetDocument = ExcelDocument.Create(targetStream)) {
            targetDocument.CopyWorksheetFrom(
                sourceDocument,
                "Data",
                "DataCopy",
                ExcelSheetNameValidationMode.Sanitize,
                new ExcelWorksheetCopyOptions { CopyMode = copyMode });
            targetDocument.Save(targetStream);
        }

        return targetStream.ToArray();
    }
}
