using System.Data.Common;
using System.Globalization;
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Reader;
#if !NET10_0_OR_GREATER
using ExcelReader.Core.ValueObjects;
#endif
using OfficeIMO.Benchmarks;
using Sylvan.Data.Excel;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Measures full typed scans on both sides of the indexed worksheet size boundary.
/// Setup verifies all values; every measurement consumes all four fields of every
/// row and checks its aggregate observation.
/// </summary>
[MemoryDiagnoser]
public class ExcelLargeTypedReadBenchmarks {
    private static readonly string[] Headers = ["Id", "Amount", "CreatedOn", "Active"];
    private string _path = string.Empty;
    private long _expected;

    [Params(25_000, 250_000, 1_000_000)]
    public int RowCount { get; set; }

    // The OfficeIMO-only first-row lane reuses this workload. Peer first-row
    // methods do not promise the same eager whole-worksheet validation contract.
    public string Operation { get; set; } = "AllRows";

    [GlobalSetup(Target = nameof(OfficeIMO))]
    public void SetupOfficeIMO() => Setup(nameof(OfficeIMO));

    [GlobalSetup(Target = nameof(Sylvan))]
    public void SetupSylvan() => Setup(nameof(Sylvan));

    [GlobalSetup(Target = nameof(ExcelReaderNet))]
    public void SetupExcelReaderNet() => Setup(nameof(ExcelReaderNet));

    /// <summary>Produces an ordinary declared dimension from a known-count row array.</summary>
    public bool DeclareDimension { get; set; }

    private void Setup(string engine) {
        string? priority = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_PROCESS_PRIORITY");
        if (!string.IsNullOrEmpty(priority)) BenchmarkProcessorAffinity.ApplyPriority(priority);
        string root = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_DATA") ?? Path.GetTempPath();
        Directory.CreateDirectory(root);
        _path = Path.Combine(root, $"large-typed-read-{Guid.NewGuid():N}.xlsx");
        try {
            using (Stream output = File.Create(_path)) {
                var rows = ExcelGeneratedRowStreamingBenchmarks.GenerateRows(RowCount);
                if (DeclareDimension) rows = rows.ToArray();
                ExcelDocument.WriteRows(output,
                    rows, Headers,
                    static (writer, row) => writer.Write(row.Id).Write(row.Amount)
                        .Write(row.CreatedOn).Write(row.Active),
                    new ExcelTabularWriteOptions { IncludeCellReferences = true, UseSharedStrings = false });
            }
            _expected = ExpectedObservation(Operation == "FirstRow" ? 1 : RowCount);
            if (engine == nameof(ExcelReaderNet)) {
                ReadExcelReader(validate: true, firstRow: false);
            } else {
                using DbDataReader reader = Open(engine);
                Read(reader, validate: true, firstRow: false);
            }
        } catch {
            Cleanup();
            throw;
        }
    }

    [GlobalCleanup]
    public void Cleanup() {
        if (_path.Length != 0) File.Delete(_path);
    }

    [Benchmark(Baseline = true)]
    public long OfficeIMO() => Measure(nameof(OfficeIMO));

    [Benchmark]
    public long Sylvan() => Measure(nameof(Sylvan));

    [Benchmark]
    public long ExcelReaderNet() => Check(ReadExcelReader(validate: false, Operation == "FirstRow"));

    private long Measure(string engine) {
        using DbDataReader reader = Open(engine);
        return Check(Read(reader, validate: false, Operation == "FirstRow"));
    }

    private DbDataReader Open(string engine) => engine == nameof(OfficeIMO)
        ? ExcelDocument.OpenDataReader(_path, new ExcelReadOptions { NumericAsDecimal = true })
        : global::Sylvan.Data.Excel.ExcelDataReader.Create(_path,
            new ExcelDataReaderOptions { Schema = ExcelSchema.Default });

    private long Read(DbDataReader reader, bool validate, bool firstRow) {
        if (reader.FieldCount != Headers.Length) throw new InvalidDataException("Incorrect field count.");
        for (int column = 0; column < Headers.Length; column++) {
            if (reader.GetName(column) != Headers[column]) throw new InvalidDataException("Incorrect header.");
        }
        long observation = 0;
        int count = 0;
        while (reader.Read()) {
            int id = reader.GetInt32(0);
            decimal amount = reader.GetDecimal(1);
            DateTime date = reader.GetDateTime(2);
            bool active = reader.GetBoolean(3);
            if (validate) ValidateRow(count, id, amount, date, active);
            Add(ref observation, id, amount, date, active);
            count++;
            if (firstRow) break;
        }
        ValidateCount(count, firstRow);
        if (!firstRow && reader.NextResult()) throw new InvalidDataException("Unexpected second sheet.");
        return observation;
    }

    private long ReadExcelReader(bool validate, bool firstRow) {
#if NET10_0_OR_GREATER
        using XlsxReader reader = ExcelReaderApi.FromXlsxFile(_path);
#else
        using XlsxReader reader = ExcelReaderApi.FromFile(_path);
#endif
        long observation = 0;
        int count = 0;
        bool header = true;
        foreach (Row row in reader) {
            if (header) {
                for (int column = 0; column < Headers.Length; column++) {
                    if (row[column].GetString() != Headers[column]) throw new InvalidDataException("Incorrect header.");
                }
                header = false;
                continue;
            }
            if (!row[0].TryParse(CultureInfo.InvariantCulture, out int id)
                || !row[1].TryParse(CultureInfo.InvariantCulture, out decimal amount)
                || !row[2].TryGetDateTime(reader.IsDate1904, out DateTime date)) {
                throw new InvalidDataException("Incorrect typed value.");
            }
            ReadOnlySpan<byte> boolean = row[3].Value;
            if (boolean.Length != 1 || boolean[0] is not ((byte)'0') and not ((byte)'1'))
                throw new InvalidDataException("Incorrect boolean value.");
            bool active = boolean[0] == (byte)'1';
            if (validate) ValidateRow(count, id, amount, date, active);
            Add(ref observation, id, amount, date, active);
            count++;
            if (firstRow) break;
        }
        ValidateCount(count, firstRow);
        return observation;
    }

    private void ValidateCount(int count, bool firstRow) {
        if (count != (firstRow ? 1 : RowCount)) throw new InvalidDataException("Incorrect row count.");
    }

    private static void ValidateRow(int index, int id, decimal amount, DateTime date, bool active) {
        var expected = ExcelGeneratedRowStreamingBenchmarks.CreateRow(index);
        if (id != expected.Id || amount != expected.Amount || date != expected.CreatedOn || active != expected.Active)
            throw new InvalidDataException($"Incorrect values at data row {index + 1}.");
    }

    private static long ExpectedObservation(int count) {
        long result = 0;
        foreach (var row in ExcelGeneratedRowStreamingBenchmarks.GenerateRows(count))
            Add(ref result, row.Id, row.Amount, row.CreatedOn, row.Active);
        return result;
    }

    private static void Add(ref long result, int id, decimal amount, DateTime date, bool active) =>
        result = unchecked(result * 31 + id + (long)(amount * 100) + date.Ticks + (active ? 1 : 0));

    private long Check(long observation) => observation == _expected
        ? observation : throw new InvalidDataException("Typed observation differs from the generated input.");
}

/// <summary>Measures the opening and eager-validation cost of OfficeIMO's first-row contract.</summary>
[MemoryDiagnoser]
public class ExcelLargeTypedFirstRowBenchmarks {
    private readonly ExcelLargeTypedReadBenchmarks _workload = new();

    [Params(25_000, 250_000, 1_000_000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        _workload.RowCount = RowCount;
        _workload.Operation = "FirstRow";
        _workload.SetupOfficeIMO();
    }

    [Benchmark]
    public long OfficeIMO() => _workload.OfficeIMO();

    [GlobalCleanup]
    public void Cleanup() => _workload.Cleanup();
}
