using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;

namespace OfficeIMO.Excel.Benchmarks;

[MemoryDiagnoser]
public class ExcelRoundTripBenchmarks {
    private byte[] _workbookBytes = [];

    [Params(250, 2500, 25000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        var rows = ExcelBenchmarkScenarioFactory.CreateSalesRecords(RowCount);
        _workbookBytes = ExcelBenchmarkScenarioFactory.CreateWorkbookBytes(rows);
        // ExcelDataReader reopens each output: every original value, the new header and the 100 edited cells.
        string detail = $"RowCount={RowCount}; ExcelDataReader reread every value and the edited column";
        ExcelOutputValidation.Check("OfficeIMO", "Load_Edit_Save", detail,
            () => ExcelSalesOutputValidator.ValidateWorkbook(OfficeIMO_Load_Edit_Save(), rows, reviewed: true));
        ExcelOutputValidation.Check("ClosedXML", "Load_Edit_Save", detail,
            () => ExcelSalesOutputValidator.ValidateWorkbook(ClosedXML_Load_Edit_Save(), rows, reviewed: true));
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("case:Load_Edit_Save")]
    public byte[] OfficeIMO_Load_Edit_Save() {
        using var input = new MemoryStream(_workbookBytes, writable: false);
        using var output = new MemoryStream();

        using (var document = ExcelDocument.Load(input)) {
            var sheet = document["Data"];
            sheet.Cell(1, 9).SetValue("ReviewStatus");

            int limit = Math.Min(RowCount + 1, 101);
            for (int row = 2; row <= limit; row++) {
                sheet.Cell(row, 9).SetValue("Reviewed");
            }

            document.Save(output);
        }

        return output.ToArray();
    }

    [Benchmark]
    [BenchmarkCategory("case:Load_Edit_Save")]
    public byte[] ClosedXML_Load_Edit_Save() {
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
        return output.ToArray();
    }
}
