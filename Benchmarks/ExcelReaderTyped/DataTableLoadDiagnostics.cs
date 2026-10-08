using System.Data;
using System.Globalization;
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Reader;
using Sylvan.Data.Excel;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks;

/// <summary>Reports native DataTable.Load contracts without pairing different result schemas.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("DataTableDiagnostics")]
public class DataTableLoadDiagnostics {
    private readonly AdoReadBenchmarks _input = new();

    [ParamsSource(nameof(RowCounts))]
    public int RowCount { get; set; } = 50_000;
    public IEnumerable<int> RowCounts() => _input.RowCounts();

    [GlobalSetup]
    public async Task SetupAsync() {
        _input.RowCount = RowCount;
        await _input.SetupAsync();
        LoadExcelReader(validate: true);
        LoadOfficeIMO(validate: true);
        LoadSylvan(validate: true);
    }

    [Benchmark]
    public int ExcelReaderDataTable() => LoadExcelReader(validate: false);
    [Benchmark]
    public int OfficeIMODataTable() => LoadOfficeIMO(validate: false);
    [Benchmark]
    public int SylvanDataTable() => LoadSylvan(validate: false);

    private int LoadExcelReader(bool validate) {
        using var workbook = ExcelReaderApi.FromXlsx(new MemoryStream(_input.WorkbookBytes, writable: false));
        using IDataReader reader = new global::ExcelReader.Core.Reader.ExcelDataReader(workbook.FirstSheet);
        return Load(reader, "ExcelReader", validate);
    }

    private int LoadOfficeIMO(bool validate) {
        using var stream = new MemoryStream(_input.WorkbookBytes, writable: false);
        using IDataReader reader = ExcelDocument.OpenDataReader(stream);
        return Load(reader, "OfficeIMO", validate);
    }

    private int LoadSylvan(bool validate) {
        using var stream = new MemoryStream(_input.WorkbookBytes, writable: false);
        using IDataReader reader = global::Sylvan.Data.Excel.ExcelDataReader.Create(stream, ExcelWorkbookType.ExcelXml);
        return Load(reader, "Sylvan", validate);
    }

    private int Load(IDataReader reader, string engine, bool validate) {
        using var table = new DataTable();
        table.Load(reader);
        if (table.Rows.Count != RowCount || table.Columns.Count != 4) throw new InvalidDataException("DataTable shape differs.");
        if (validate) {
            Console.WriteLine($"{engine} DataTable schema: {string.Join(", ", table.Columns.Cast<DataColumn>().Select(column => $"{column.ColumnName}:{column.DataType.Name}"))}.");
            // Sylvan's string schema exposes invariant XML lexical values; DataTable
            // converts ExcelReader's native values to strings using current culture.
            CultureInfo valueCulture = engine == "Sylvan" ? CultureInfo.InvariantCulture : CultureInfo.CurrentCulture;
            for (int index = 0; index < table.Rows.Count; index++) {
                DataRow row = table.Rows[index];
                object rawDate = row[2];
                DateTime date = rawDate is DateTime nativeDate ? nativeDate
                    : rawDate is string text ? DateTime.Parse(text, valueCulture)
                    : DateTime.FromOADate(Convert.ToDouble(rawDate, CultureInfo.InvariantCulture));
                TypedWorkbookFixture.ValidateRecord(new TypedRecord {
                    Name = Convert.ToString(row[0], valueCulture),
                    Id = Convert.ToInt32(row[1], valueCulture), Date = date,
                    Value = Convert.ToDouble(row[3], valueCulture)
                }, index + 1);
            }
        }
        return table.Rows.Count;
    }
}
