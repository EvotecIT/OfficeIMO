using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using ClosedXML.Excel;

namespace OfficeIMO.Excel.Benchmarks;

[MemoryDiagnoser]
[SimpleJob(RuntimeMoniker.Net80)]
public class ExcelWriteBenchmarks {
    private IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> _rows = null!;

    [Params(250, 2500)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        _rows = ExcelBenchmarkScenarioFactory.CreateSalesRecords(RowCount);
    }

    [Benchmark(Baseline = true)]
    public int OfficeIMO_Write_Report() {
        using var stream = new MemoryStream();

        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Data");
            ExcelBenchmarkScenarioFactory.PopulateOfficeImoWorksheet(sheet, _rows);
        }

        return checked((int)stream.Length);
    }

    [Benchmark]
    public int ClosedXML_Write_Report() {
        using var stream = new MemoryStream();
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Data");

        ExcelBenchmarkScenarioFactory.PopulateClosedXmlWorksheet(worksheet, _rows);
        workbook.SaveAs(stream);

        return checked((int)stream.Length);
    }
}
