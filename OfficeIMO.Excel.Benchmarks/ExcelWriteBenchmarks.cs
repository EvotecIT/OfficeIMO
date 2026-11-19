using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;

namespace OfficeIMO.Excel.Benchmarks;

[MemoryDiagnoser]
public class ExcelWriteBenchmarks {
    private IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> _rows = null!;

    [Params(250, 2500, 25000, 100_000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        _rows = ExcelBenchmarkScenarioFactory.CreateSalesRecords(RowCount);
        // ExcelDataReader reopens each output and checks the sheet, the headers and every value.
        string detail = $"RowCount={RowCount}; ExcelDataReader reread every value";
        ExcelOutputValidation.Check("OfficeIMO", "Write_Report", detail,
            () => ExcelSalesOutputValidator.ValidateWorkbook(OfficeIMO_Write_Report(), _rows));
        ExcelOutputValidation.Check("ClosedXML", "Write_Report", detail,
            () => ExcelSalesOutputValidator.ValidateWorkbook(ClosedXML_Write_Report(), _rows));
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("case:Write_Report")]
    public byte[] OfficeIMO_Write_Report() {
        using var stream = new MemoryStream();

        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddSheet("Data");
            ExcelBenchmarkScenarioFactory.PopulateOfficeImoWorksheet(sheet, _rows);
            document.Save(stream);
        }

        return stream.ToArray();
    }

    [Benchmark]
    [BenchmarkCategory("case:Write_Report")]
    public byte[] ClosedXML_Write_Report() {
        using var stream = new MemoryStream();
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Data");

        ExcelBenchmarkScenarioFactory.PopulateClosedXmlWorksheet(worksheet, _rows);
        workbook.SaveAs(stream);

        return stream.ToArray();
    }
}
