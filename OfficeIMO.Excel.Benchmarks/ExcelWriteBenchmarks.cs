using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;

namespace OfficeIMO.Excel.Benchmarks;

[MemoryDiagnoser]
public class ExcelWriteBenchmarks {
    private IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> _rows = null!;

    [Params(250, 2500, 25000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        _rows = ExcelBenchmarkScenarioFactory.CreateSalesRecords(RowCount);
        ExcelSalesOutputValidator.ValidateWorkbook(OfficeIMO_Write_Report(), _rows);
        ExcelSalesOutputValidator.ValidateWorkbook(ClosedXML_Write_Report(), _rows);
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO_Write_Report() {
        using var stream = new MemoryStream();

        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorksheet("Data");
            ExcelBenchmarkScenarioFactory.PopulateOfficeImoWorksheet(sheet, _rows);
            document.Save(stream);
        }

        return stream.ToArray();
    }

    [Benchmark]
    public byte[] ClosedXML_Write_Report() {
        using var stream = new MemoryStream();
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Data");

        ExcelBenchmarkScenarioFactory.PopulateClosedXmlWorksheet(worksheet, _rows);
        workbook.SaveAs(stream);

        return stream.ToArray();
    }
}
