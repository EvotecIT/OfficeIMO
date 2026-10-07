using System.Data;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using ClosedXML.Excel;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Reads every field into the same materialized result shape in each category.
/// Workbook opening, decoding, and materialization are timed; setup checks every value.
/// </summary>
[MemoryDiagnoser]
[GroupBenchmarksBy(BenchmarkLogicalGroupRule.ByCategory)]
[CategoriesColumn]
public class ExcelReadBenchmarks {
    private byte[] _workbookBytes = [];
    private string _range = string.Empty;

    [Params(250, 2500, 25000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        var rows = ExcelBenchmarkScenarioFactory.CreateSalesRecords(RowCount);
        _workbookBytes = ExcelBenchmarkScenarioFactory.CreateWorkbookBytes(rows);
        _range = ExcelBenchmarkScenarioFactory.BuildDataRange(RowCount);
        ExcelSalesOutputValidator.ValidateDictionaries(rows, OfficeIMO_Read_Objects());
        ExcelSalesOutputValidator.ValidateDictionaries(rows, ClosedXML_Read_Objects());
        using DataTable office = OfficeIMO_Read_DataTable();
        using DataTable closed = ClosedXML_Read_DataTable();
        ExcelSalesOutputValidator.ValidateDataTable(rows, office);
        ExcelSalesOutputValidator.ValidateDataTable(rows, closed);
    }

    [Benchmark(Baseline = true), BenchmarkCategory("Dictionaries")]
    public List<Dictionary<string, object?>> OfficeIMO_Read_Objects() {
        using var stream = new MemoryStream(_workbookBytes, writable: false);
        using var reader = ExcelDocumentReader.Open(stream);
        return reader.GetSheet("Data").ReadObjects(_range).ToList();
    }

    [Benchmark(Baseline = true), BenchmarkCategory("DataTable")]
    public DataTable OfficeIMO_Read_DataTable() {
        using var stream = new MemoryStream(_workbookBytes, writable: false);
        using var reader = ExcelDocumentReader.Open(stream, new ExcelReadOptions { InferDataTableColumnTypes = false });
        return reader.GetSheet("Data").ReadRangeAsDataTable(_range, headersInFirstRow: true);
    }

    [Benchmark(Baseline = true), BenchmarkCategory("DataTableCount")]
    public int OfficeIMO_Read_DataTableCount() {
        using DataTable table = OfficeIMO_Read_DataTable();
        return CheckTableShape(table);
    }

    [Benchmark, BenchmarkCategory("DataTableCount")]
    public int ClosedXML_Read_DataTableCount() {
        using DataTable table = ClosedXML_Read_DataTable();
        return CheckTableShape(table);
    }

    [Benchmark, BenchmarkCategory("Dictionaries")]
    public List<Dictionary<string, object?>> ClosedXML_Read_Objects() {
        using var stream = new MemoryStream(_workbookBytes, writable: false);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheet("Data");
        string[] headers = ReadHeaders(worksheet);
        var rows = new List<Dictionary<string, object?>>(RowCount);
        for (int row = 2; row <= RowCount + 1; row++) {
            var values = new Dictionary<string, object?>(headers.Length, StringComparer.OrdinalIgnoreCase);
            for (int column = 0; column < headers.Length; column++) {
                values.Add(headers[column], ReadValue(worksheet.Cell(row, column + 1)));
            }
            rows.Add(values);
        }
        return rows;
    }

    [Benchmark, BenchmarkCategory("DataTable")]
    public DataTable ClosedXML_Read_DataTable() {
        using var stream = new MemoryStream(_workbookBytes, writable: false);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheet("Data");
        string[] headers = ReadHeaders(worksheet);
        var table = new DataTable { MinimumCapacity = RowCount };
        foreach (string header in headers) table.Columns.Add(header, typeof(object));
        table.BeginLoadData();
        try {
            var values = new object?[headers.Length];
            for (int row = 2; row <= RowCount + 1; row++) {
                for (int column = 0; column < headers.Length; column++) {
                    values[column] = ReadValue(worksheet.Cell(row, column + 1));
                }
                table.Rows.Add(values);
            }
        } finally {
            table.EndLoadData();
        }
        return table;
    }

    private static string[] ReadHeaders(IXLWorksheet worksheet) =>
        Enumerable.Range(1, ExcelBenchmarkScenarioFactory.SalesColumnNames.Length)
            .Select(column => worksheet.Cell(1, column).GetString()).ToArray();

    private int CheckTableShape(DataTable table) =>
        table.Rows.Count == RowCount && table.Columns.Count == ExcelBenchmarkScenarioFactory.SalesColumnNames.Length
            ? table.Rows.Count : throw new InvalidDataException("DataTable shape differs from the validated fixture.");

    private static object? ReadValue(IXLCell cell) => cell.DataType switch {
        XLDataType.Blank => null,
        XLDataType.Boolean => cell.GetBoolean(),
        XLDataType.Number => cell.GetDouble(),
        XLDataType.DateTime => cell.GetDateTime(),
        XLDataType.Text => cell.GetString(),
        _ => throw new InvalidDataException($"Unexpected fixture cell type {cell.DataType}.")
    };
}
