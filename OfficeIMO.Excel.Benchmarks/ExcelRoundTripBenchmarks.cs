using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using ClosedXML.Excel;

namespace OfficeIMO.Excel.Benchmarks;

[MemoryDiagnoser]
[SimpleJob(RuntimeMoniker.Net80)]
public class ExcelRoundTripBenchmarks {
    private byte[] _workbookBytes = [];

    [Params(250, 2500)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        var rows = ExcelBenchmarkScenarioFactory.CreateSalesRecords(RowCount);
        _workbookBytes = ExcelBenchmarkScenarioFactory.CreateWorkbookBytes(rows);
    }

    [Benchmark(Baseline = true)]
    public int OfficeIMO_Load_Edit_Save() {
        using var input = new MemoryStream(_workbookBytes, writable: false);
        using var output = new MemoryStream();

        using (var document = ExcelDocument.Load(input)) {
            var sheet = document["Data"];
            sheet.CellValue(1, 9, "ReviewStatus");

            int limit = Math.Min(RowCount + 1, 101);
            for (int row = 2; row <= limit; row++) {
                sheet.CellValue(row, 9, "Reviewed");
            }

            document.Save(output);
        }

        return checked((int)output.Length);
    }

    [Benchmark]
    public int ClosedXML_Load_Edit_Save() {
        using var input = new MemoryStream(_workbookBytes, writable: false);
        using var output = new MemoryStream();
        using var workbook = new XLWorkbook(input);
        var worksheet = workbook.Worksheet("Data");

        worksheet.Cell(1, 9).Value = "ReviewStatus";
        int limit = Math.Min(RowCount + 1, 101);
        for (int row = 2; row <= limit; row++) {
            worksheet.Cell(row, 9).Value = "Reviewed";
        }

        workbook.SaveAs(output);
        return checked((int)output.Length);
    }
}
