using System.Data;
using System.Data.Common;
using System.Globalization;
using System.Text;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using ExcelDataReader;
using OfficeIMO.Benchmarks;
using OfficeIMO.Data;
using Sylvan.Data;
using Sylvan.Data.Excel;
using MiniExcelApi = MiniExcelLibs.MiniExcel;
using SylvanExcelDataReader = Sylvan.Data.Excel.ExcelDataReader;
#if NET10_0_OR_GREATER
using ExcelReader.Core.Parser;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;
#endif

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Maps every row of the hash-pinned 65K sales XLSX into typed objects or a DataTable. Every lane
/// converts all fourteen columns and must reproduce the observation computed from the matching CSV
/// fixture; setup runs each lane once before measurement.
/// </summary>
[MemoryDiagnoser]
[GroupBenchmarksBy(BenchmarkLogicalGroupRule.ByCategory)]
[CategoriesColumn]
public class MarkPflug65KTypedMappingXlsxBenchmarks {
    private string _path = string.Empty;
    private long _expected;
#if NET10_0_OR_GREATER
    private readonly ExcelParser<SalesRow> _excelReaderParser = ExcelParser.FromAttributes<SalesRow>();
#endif

    [GlobalSetup]
    public void Setup() {
        BenchmarkProcessPriority.ApplyConfiguredPriority();
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        _expected = SalesMappingFixture.ExpectedObservation(SalesMappingFixture.CreateCsv());
        _path = SalesMappingFixture.CreateXlsx(Path.Combine(MarkPflug65KFixture.Root, "typed-mapping"));
        OfficeIMO_Class(); OfficeIMO_ClassParallel(); Sylvan_Class(); MiniExcel_Class();
        OfficeIMO_Record();
        OfficeIMO_DataTable(); ExcelDataReader_DataSet(); Sylvan_DataTable();
#if NET10_0_OR_GREATER
        ExcelReaderNet_Class();
#endif
    }

    private static ExcelReaderOptions OfficeOptions => new() { NumericAsDecimal = true };

    [Benchmark(Baseline = true), BenchmarkCategory("Class")]
    public long OfficeIMO_Class() {
        using DbDataReader reader = ExcelDocument.OpenReader(_path, OfficeOptions);
        return Observe(reader.RowsAs<SalesRow>(), nameof(OfficeIMO_Class));
    }

    [Benchmark, BenchmarkCategory("Class")]
    public long OfficeIMO_ClassParallel() {
        using DbDataReader reader = ExcelDocument.OpenReader(_path, OfficeOptions);
        return Observe(reader.RowsAsParallel<SalesRow>(), nameof(OfficeIMO_ClassParallel));
    }

    [Benchmark, BenchmarkCategory("Class")]
    public long Sylvan_Class() {
        using SylvanExcelDataReader reader = SylvanExcelDataReader.Create(_path,
            new ExcelDataReaderOptions { Schema = ExcelSchema.Default });
        return Observe(reader.GetRecords<SalesRow>(), nameof(Sylvan_Class));
    }

    [Benchmark, BenchmarkCategory("Class")]
    public long MiniExcel_Class() =>
        Observe(MiniExcelApi.Query<SalesRow>(_path), nameof(MiniExcel_Class));

#if NET10_0_OR_GREATER
    [Benchmark, BenchmarkCategory("Class")]
    public long ExcelReaderNet_Class() {
        using var reader = ExcelReaderApi.FromXlsxFile(_path);
        return Observe(_excelReaderParser.Parse(reader.FirstSheet), nameof(ExcelReaderNet_Class));
    }
#endif

    [Benchmark(Baseline = true), BenchmarkCategory("Record")]
    public long OfficeIMO_Record() {
        using DbDataReader reader = ExcelDocument.OpenReader(_path, OfficeOptions);
        return Observe(reader.RowsAs<SalesRecord>(), nameof(OfficeIMO_Record));
    }

    /// <summary>Fills a DataTable with each library's own cell types; values are checked after conversion.</summary>
    [Benchmark(Baseline = true), BenchmarkCategory("DataTable")]
    public long OfficeIMO_DataTable() {
        using DbDataReader reader = ExcelDocument.OpenReader(_path, OfficeOptions);
        var table = new DataTable();
        table.Load(reader);
        return ObserveTable(table, nameof(OfficeIMO_DataTable));
    }

    [Benchmark, BenchmarkCategory("DataTable")]
    public long ExcelDataReader_DataSet() {
        using FileStream stream = File.OpenRead(_path);
        using IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
        DataSet set = reader.AsDataSet(new ExcelDataSetConfiguration {
            ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = true }
        });
        return ObserveTable(set.Tables[0], nameof(ExcelDataReader_DataSet));
    }

    [Benchmark, BenchmarkCategory("DataTable")]
    public long Sylvan_DataTable() {
        using SylvanExcelDataReader reader = SylvanExcelDataReader.Create(_path,
            new ExcelDataReaderOptions { Schema = ExcelSchema.Default });
        var table = new DataTable();
        table.Load(reader);
        return ObserveTable(table, nameof(Sylvan_DataTable));
    }

    private long ObserveTable(DataTable table, string lane) {
        for (int index = 0; index < SalesMappingFixture.Headers.Length; index++) {
            if (table.Columns[index].ColumnName != SalesMappingFixture.Headers[index])
                throw new InvalidDataException($"{lane} named column {index} '{table.Columns[index].ColumnName}'.");
        }

        var observation = new SalesObservation();
        foreach (DataRow row in table.Rows) {
            object?[] v = row.ItemArray;
            observation.Add(Text(v[0]), Text(v[1]), Text(v[2]), Text(v[3]), Text(v[4]),
                Date(v[5]), Convert.ToInt32(v[6], CultureInfo.InvariantCulture), Date(v[7]),
                Convert.ToInt32(v[8], CultureInfo.InvariantCulture),
                Number(v[9]), Number(v[10]), Number(v[11]), Number(v[12]), Number(v[13]));
        }

        return observation.Check(_expected, lane);
    }

    private static string Text(object? value) => Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;

    private static DateTime Date(object? value) => value switch {
        DateTime date => date,
        double serial => DateTime.FromOADate(serial),
        decimal serial => DateTime.FromOADate((double)serial),
        string text => DateTime.Parse(text, CultureInfo.InvariantCulture),
        _ => throw new InvalidDataException($"Unexpected date cell {value?.GetType().Name} '{value}'.")
    };

    private static decimal Number(object? value) => Convert.ToDecimal(value, CultureInfo.InvariantCulture);

    private long Observe(IEnumerable<SalesRow> rows, string lane) {
        var observation = new SalesObservation();
        foreach (SalesRow row in rows) observation.Add(row);
        return observation.Check(_expected, lane);
    }

    private long Observe(IEnumerable<SalesRecord> rows, string lane) {
        var observation = new SalesObservation();
        foreach (SalesRecord row in rows) observation.Add(row);
        return observation.Check(_expected, lane);
    }
}
