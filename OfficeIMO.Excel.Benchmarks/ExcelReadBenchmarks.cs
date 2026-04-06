using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using ClosedXML.Excel;

namespace OfficeIMO.Excel.Benchmarks;

[MemoryDiagnoser]
[SimpleJob(RuntimeMoniker.Net80)]
public class ExcelReadBenchmarks {
    private byte[] _workbookBytes = [];
    private string _range = string.Empty;

    [Params(250, 2500)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        var rows = ExcelBenchmarkScenarioFactory.CreateSalesRecords(RowCount);
        _workbookBytes = ExcelBenchmarkScenarioFactory.CreateWorkbookBytes(rows);
        _range = ExcelBenchmarkScenarioFactory.BuildDataRange(RowCount);
    }

    [Benchmark(Baseline = true)]
    public int OfficeIMO_Read_Objects() {
        using var stream = new MemoryStream(_workbookBytes, writable: false);
        using var reader = ExcelDocumentReader.Open(stream);
        return reader.GetSheet("Data").ReadObjects(_range).Count();
    }

    [Benchmark]
    public int OfficeIMO_Read_DataTable() {
        using var stream = new MemoryStream(_workbookBytes, writable: false);
        using var reader = ExcelDocumentReader.Open(stream);
        return reader.GetSheet("Data").ReadRangeAsDataTable(_range, headersInFirstRow: true).Rows.Count;
    }

    [Benchmark]
    public int ClosedXML_Read_Rows() {
        using var stream = new MemoryStream(_workbookBytes, writable: false);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheet("Data");
        int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        int count = 0;

        for (int row = 2; row <= lastRow; row++) {
            _ = worksheet.Cell(row, 1).GetValue<int>();
            _ = worksheet.Cell(row, 5).GetValue<double>();
            count++;
        }

        return count;
    }
}
