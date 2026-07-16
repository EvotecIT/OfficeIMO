using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;

namespace OfficeIMO.Excel.Benchmarks;

[MemoryDiagnoser]
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
            var sheet = document.AddWorksheet("Data");
            ExcelBenchmarkScenarioFactory.PopulateOfficeImoWorksheet(sheet, _rows);
            document.Save(stream);
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
