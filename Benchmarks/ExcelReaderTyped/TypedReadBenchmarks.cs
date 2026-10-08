using System.Data.Common;
using System.Globalization;
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using Sylvan.Data;
using Sylvan.Data.Excel;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks;

/// <summary>Equivalent materialized typed scans of the independent producer's workbook.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("Typed")]
public class TypedReadBenchmarks {
    private byte[] _workbook = [];
    private long _expected;
    internal byte[] Workbook => _workbook;

    [ParamsSource(nameof(RowCounts))]
    public int RowCount { get; set; } = 50_000;

    [ParamsSource(nameof(Shapes))]
    public string Shape { get; set; } = "Original";

    public IEnumerable<int> RowCounts() {
        string value = Environment.GetEnvironmentVariable("OFFICEIMO_TYPED_BENCHMARK_ROWS") ?? "50000";
        foreach (string item in value.Split(',', StringSplitOptions.TrimEntries)) {
            if (!int.TryParse(item, NumberStyles.None, CultureInfo.InvariantCulture, out int count) || count is < 1 or > 1_000_000)
                throw new ArgumentException("OFFICEIMO_TYPED_BENCHMARK_ROWS requires counts between 1 and 1000000.");
            yield return count;
        }
    }

    public IEnumerable<string> Shapes() {
        string value = Environment.GetEnvironmentVariable("OFFICEIMO_TYPED_BENCHMARK_SHAPES") ?? "Original";
        foreach (string shape in value.Split(',', StringSplitOptions.TrimEntries)) {
            if (shape is not ("Original" or "SharedStrings" or "RowReferences" or "Dimension"))
                throw new ArgumentException("OFFICEIMO_TYPED_BENCHMARK_SHAPES requires Original, SharedStrings, RowReferences, or Dimension.");
            yield return shape;
        }
    }

    [GlobalSetup]
    public async Task SetupAsync() {
        BenchmarkInput.WriteDescription();
        _workbook = await TypedWorkbookFixture.CreateAsync(RowCount, Shape);
        _expected = TypedWorkbookFixture.ExpectedChecksum(RowCount);
        string description = TypedWorkbookFixture.ValidateShape(_workbook, RowCount, Shape);
        ValidateExcelReader();
        ValidateDataReader(useSylvan: false);
        ValidateDataReader(useSylvan: true);
        Check(ExcelReaderTyped(), RowCount);
        Check(OfficeIMOTyped(), RowCount);
        Check(SylvanTyped(), RowCount);
        Console.WriteLine($"Validated {Shape}: {description}, checksum={_expected}.");
    }

    [Benchmark(Baseline = true)]
    public long ExcelReaderTyped() {
        using var stream = new MemoryStream(_workbook, writable: false);
        using var workbook = ExcelReaderApi.FromXlsx(stream);
        long sum = 0;
        int count = 0;
        foreach (TypedRecord record in ExcelParser.FromAttributes<TypedRecord>().Parse(workbook.FirstSheet)) {
            sum = unchecked(sum + TypedWorkbookFixture.Accumulate(record));
            count++;
        }
        return Check(sum, count);
    }

    [Benchmark]
    public long OfficeIMOTyped() {
        using var reader = ExcelDocument.OpenDataReader(_workbook, new ExcelReadOptions { HasHeaderRow = true });
        long sum = 0;
        int count = 0;
        while (reader.Read()) {
            sum = unchecked(sum + TypedWorkbookFixture.Accumulate(new TypedRecord {
                Name = reader.GetString(0), Id = reader.GetInt32(1),
                Date = reader.GetDateTime(2), Value = reader.GetDouble(3),
            }));
            count++;
        }
        return Check(sum, count);
    }

    [Benchmark]
    public long SylvanTyped() {
        using var stream = new MemoryStream(_workbook, writable: false);
        using var reader = global::Sylvan.Data.Excel.ExcelDataReader.Create(stream,
            ExcelWorkbookType.ExcelXml, new ExcelDataReaderOptions());
        long sum = 0;
        int count = 0;
        foreach (TypedRecord record in reader.GetRecords<TypedRecord>()) {
            sum = unchecked(sum + TypedWorkbookFixture.Accumulate(record));
            count++;
        }
        return Check(sum, count);
    }

    internal int OpenOfficeIMO() {
        using var reader = ExcelDocument.OpenDataReader(_workbook, new ExcelReadOptions { HasHeaderRow = true });
        return reader.FieldCount == TypedWorkbookFixture.Headers.Length
            ? reader.FieldCount : throw new InvalidDataException("Incorrect field count when opening the workbook.");
    }

    private void ValidateExcelReader() {
        using var stream = new MemoryStream(_workbook, writable: false);
        using var workbook = ExcelReaderApi.FromXlsx(stream);
        int count = 0;
        foreach (TypedRecord record in ExcelParser.FromAttributes<TypedRecord>().Parse(workbook.FirstSheet))
            TypedWorkbookFixture.ValidateRecord(record, ++count);
        if (count != RowCount) throw new InvalidDataException("ExcelReader returned an incorrect row count.");
    }

    private void ValidateDataReader(bool useSylvan) {
        using var stream = new MemoryStream(_workbook, writable: false);
        using DbDataReader reader = useSylvan
            ? global::Sylvan.Data.Excel.ExcelDataReader.Create(stream, ExcelWorkbookType.ExcelXml, new ExcelDataReaderOptions())
            : ExcelDocument.OpenDataReader(_workbook, new ExcelReadOptions { HasHeaderRow = true });
        if (reader.FieldCount != TypedWorkbookFixture.Headers.Length) throw new InvalidDataException("Incorrect field count.");
        for (int column = 0; column < TypedWorkbookFixture.Headers.Length; column++) {
            if (reader.GetName(column) != TypedWorkbookFixture.Headers[column])
                throw new InvalidDataException("Incorrect header.");
        }
        int count = 0;
        while (reader.Read()) TypedWorkbookFixture.ValidateRecord(new TypedRecord {
            Name = reader.GetString(0), Id = reader.GetInt32(1),
            Date = reader.GetDateTime(2), Value = reader.GetDouble(3),
        }, ++count);
        if (count != RowCount || reader.NextResult()) throw new InvalidDataException("Incorrect row or sheet count.");
    }

    internal long Check(long checksum, int count) => checksum == _expected && count == RowCount
        ? checksum : throw new InvalidDataException("Typed scan returned an incorrect count or checksum.");
}

/// <summary>Separately measures OfficeIMO reader opening and eager worksheet qualification.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("OpenOnly")]
public class ReaderOpeningBenchmarks {
    private readonly TypedReadBenchmarks _workload = new();

    [ParamsSource(nameof(RowCounts))]
    public int RowCount { get; set; } = 50_000;

    [ParamsSource(nameof(Shapes))]
    public string Shape { get; set; } = "Original";

    public IEnumerable<int> RowCounts() => _workload.RowCounts();
    public IEnumerable<string> Shapes() => _workload.Shapes();

    [GlobalSetup]
    public Task SetupAsync() {
        _workload.RowCount = RowCount;
        _workload.Shape = Shape;
        return _workload.SetupAsync();
    }

    [Benchmark]
    public int OfficeIMOOpenOnly() => _workload.OpenOfficeIMO();
}
