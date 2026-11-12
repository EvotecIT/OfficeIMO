using System.Data;
using System.Data.Common;
using System.Globalization;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using nietras.SeparatedValues;
using OfficeIMO.Benchmarks;
using OfficeIMO.Data;
using Sylvan.Data;
using SepLib = nietras.SeparatedValues.Sep;
using SylvanCsvDataReader = Sylvan.Data.Csv.CsvDataReader;
#if NET10_0_OR_GREATER
using ExcelReader.Core.Parser;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;
#endif

namespace OfficeIMO.CSV.Benchmarks;

[global::CsvHelper.CsvRecordMap(typeof(SalesRow))]
internal partial class SalesRowCsvHelperMap { }

[global::CsvHelper.CsvRecordMap(typeof(SalesRecord))]
internal partial class SalesRecordCsvHelperMap { }

/// <summary>
/// Maps every row of the hash-pinned 65K sales CSV into typed objects or a DataTable. Every lane
/// converts all fourteen columns with invariant culture and must reproduce the observation computed
/// by splitting the text directly; setup runs each lane once before measurement.
/// </summary>
[MemoryDiagnoser]
[GroupBenchmarksBy(BenchmarkLogicalGroupRule.ByCategory)]
[CategoriesColumn]
public class MarkPflug65KTypedMappingCsvBenchmarks {
    private byte[] _csv = [];
    private long _expected;
#if NET10_0_OR_GREATER
    private readonly ExcelParser<SalesRow> _excelReaderParser = ExcelParser.FromAttributes<SalesRow>();
#endif

    [GlobalSetup]
    public void Setup() {
        BenchmarkProcessPriority.ApplyConfiguredPriority();
        _csv = SalesMappingFixture.CreateCsv();
        _expected = SalesMappingFixture.ExpectedObservation(_csv);
        OfficeIMO_Class(); OfficeIMO_ClassParallel(); CsvHelper_Class(); Sylvan_Class(); Sep_Class();
        OfficeIMO_Record(); CsvHelper_Record(); Sep_Record();
        OfficeIMO_DataTable(); Sylvan_DataTable();
#if NET10_0_OR_GREATER
        ExcelReaderNet_Class();
#endif
    }

    private MemoryStream Input() => new(_csv, writable: false);

    private static CsvReaderOptions OfficeOptions => new() { Culture = CultureInfo.InvariantCulture, DetectDelimiter = false };

    [Benchmark(Baseline = true), BenchmarkCategory("Class")]
    public long OfficeIMO_Class() {
        using DbDataReader reader = CsvDocument.OpenReader(Input(), OfficeOptions);
        return Observe(reader.RowsAs<SalesRow>(), nameof(OfficeIMO_Class));
    }

    [Benchmark, BenchmarkCategory("Class")]
    public long OfficeIMO_ClassParallel() {
        using DbDataReader reader = CsvDocument.OpenReader(Input(), OfficeOptions);
        return Observe(reader.RowsAsParallel<SalesRow>(), nameof(OfficeIMO_ClassParallel));
    }

    [Benchmark, BenchmarkCategory("Class")]
    public long CsvHelper_Class() {
        using var reader = global::CsvHelper.CsvReader.Create(Input(),
            new global::CsvHelper.CsvReaderOptions { FormatProvider = CultureInfo.InvariantCulture });
        return Observe(reader.GetRecords<SalesRow>(), nameof(CsvHelper_Class));
    }

    [Benchmark, BenchmarkCategory("Class")]
    public long Sylvan_Class() {
        using var text = new StreamReader(Input());
        using SylvanCsvDataReader reader = SylvanCsvDataReader.Create(text);
        return Observe(reader.GetRecords<SalesRow>(), nameof(Sylvan_Class));
    }

    [Benchmark, BenchmarkCategory("Class")]
    public long Sep_Class() {
        using SepReader reader = SepLib.Reader(o => o with { CultureInfo = CultureInfo.InvariantCulture }).From(_csv);
        var observation = new SalesObservation();
        foreach (SepReader.Row row in reader) {
            observation.Add(new SalesRow {
                Region = row[0].ToString(), Country = row[1].ToString(), ItemType = row[2].ToString(),
                SalesChannel = row[3].ToString(), OrderPriority = row[4].ToString(), OrderDate = row[5].Parse<DateTime>(),
                OrderId = row[6].Parse<int>(), ShipDate = row[7].Parse<DateTime>(), UnitsSold = row[8].Parse<int>(),
                UnitPrice = row[9].Parse<decimal>(), UnitCost = row[10].Parse<decimal>(), TotalRevenue = row[11].Parse<decimal>(),
                TotalCost = row[12].Parse<decimal>(), TotalProfit = row[13].Parse<decimal>()
            });
        }

        return observation.Check(_expected, nameof(Sep_Class));
    }

#if NET10_0_OR_GREATER
    [Benchmark, BenchmarkCategory("Class")]
    public long ExcelReaderNet_Class() {
        using var reader = ExcelReaderApi.FromCsv(_csv);
        return Observe(_excelReaderParser.Parse(reader.FirstSheet), nameof(ExcelReaderNet_Class));
    }
#endif

    [Benchmark(Baseline = true), BenchmarkCategory("Record")]
    public long OfficeIMO_Record() {
        using DbDataReader reader = CsvDocument.OpenReader(Input(), OfficeOptions);
        return Observe(reader.RowsAs<SalesRecord>(), nameof(OfficeIMO_Record));
    }

    [Benchmark, BenchmarkCategory("Record")]
    public long CsvHelper_Record() {
        using var reader = global::CsvHelper.CsvReader.Create(Input(),
            new global::CsvHelper.CsvReaderOptions { FormatProvider = CultureInfo.InvariantCulture });
        return Observe(reader.GetRecords<SalesRecord>(), nameof(CsvHelper_Record));
    }


    [Benchmark, BenchmarkCategory("Record")]
    public long Sep_Record() {
        using SepReader reader = SepLib.Reader(o => o with { CultureInfo = CultureInfo.InvariantCulture }).From(_csv);
        var observation = new SalesObservation();
        foreach (SepReader.Row row in reader) {
            observation.Add(new SalesRecord(
                row[0].ToString(), row[1].ToString(), row[2].ToString(), row[3].ToString(), row[4].ToString(),
                row[5].Parse<DateTime>(), row[6].Parse<int>(), row[7].Parse<DateTime>(), row[8].Parse<int>(),
                row[9].Parse<decimal>(), row[10].Parse<decimal>(), row[11].Parse<decimal>(), row[12].Parse<decimal>(),
                row[13].Parse<decimal>()));
        }

        return observation.Check(_expected, nameof(Sep_Record));
    }

    /// <summary>Loads a DataTable whose fourteen columns carry the mapped CLR types.</summary>
    [Benchmark(Baseline = true), BenchmarkCategory("DataTable")]
    public long OfficeIMO_DataTable() {
        using DbDataReader reader = CsvDocument.OpenReader(Input(), new CsvReaderOptions {
            Culture = CultureInfo.InvariantCulture,
            DetectDelimiter = false,
            Schema = BuildOfficeSchema()
        });
        return ObserveTable(LoadTable(reader), nameof(OfficeIMO_DataTable));
    }

    [Benchmark, BenchmarkCategory("DataTable")]
    public long Sylvan_DataTable() {
        using var text = new StreamReader(Input());
        using SylvanCsvDataReader reader = SylvanCsvDataReader.Create(text,
            new Sylvan.Data.Csv.CsvDataReaderOptions { Schema = new Sylvan.Data.Csv.CsvSchema(SylvanSchema()) });
        return ObserveTable(LoadTable(reader), nameof(Sylvan_DataTable));
    }

    private static DataTable LoadTable(DbDataReader reader) {
        var table = new DataTable();
        table.Load(reader);
        return table;
    }

    private static readonly Type[] ColumnTypes = [
        typeof(string), typeof(string), typeof(string), typeof(string), typeof(string), typeof(DateTime), typeof(int),
        typeof(DateTime), typeof(int), typeof(decimal), typeof(decimal), typeof(decimal), typeof(decimal), typeof(decimal)
    ];

    private static CsvSchema BuildOfficeSchema() {
        var schema = new CsvSchemaBuilder();
        for (int index = 0; index < SalesMappingFixture.Headers.Length; index++)
            schema.Column(SalesMappingFixture.Headers[index]).AsType(ColumnTypes[index]);
        return schema.Build();
    }

    private static Schema SylvanSchema() {
        string[] h = SalesMappingFixture.Headers;
        return new Schema.Builder()
            .Add<string>(h[0]).Add<string>(h[1]).Add<string>(h[2]).Add<string>(h[3]).Add<string>(h[4])
            .Add<DateTime>(h[5]).Add<int>(h[6]).Add<DateTime>(h[7]).Add<int>(h[8])
            .Add<decimal>(h[9]).Add<decimal>(h[10]).Add<decimal>(h[11]).Add<decimal>(h[12]).Add<decimal>(h[13])
            .Build();
    }

    private long ObserveTable(DataTable table, string lane) {
        for (int index = 0; index < ColumnTypes.Length; index++) {
            if (table.Columns[index].DataType != ColumnTypes[index])
                throw new InvalidDataException($"{lane} typed column {index} as {table.Columns[index].DataType}.");
        }

        var observation = new SalesObservation();
        foreach (DataRow row in table.Rows) {
            object?[] v = row.ItemArray;
            observation.Add((string)v[0]!, (string)v[1]!, (string)v[2]!, (string)v[3]!, (string)v[4]!,
                (DateTime)v[5]!, (int)v[6]!, (DateTime)v[7]!, (int)v[8]!,
                (decimal)v[9]!, (decimal)v[10]!, (decimal)v[11]!, (decimal)v[12]!, (decimal)v[13]!);
        }

        return observation.Check(_expected, lane);
    }

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
