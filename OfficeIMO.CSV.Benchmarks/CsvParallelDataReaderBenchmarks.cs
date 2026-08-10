using System.Globalization;
using System.Text;
using BenchmarkDotNet.Attributes;
using Dataplat.Dbatools.Csv.Reader;
using OfficeIMO.CSV;
using DataplatCsvDataReader = Dataplat.Dbatools.Csv.Reader.CsvDataReader;
using DataplatCsvReaderOptions = Dataplat.Dbatools.Csv.Reader.CsvReaderOptions;

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>
/// Equivalent-contract typed data-reader comparison for SQL bulk-copy-shaped consumption.
/// Every lane opens the same UTF-8 file, converts the same five columns, reads every value,
/// preserves source order, and returns the same checksum.
/// </summary>
[MemoryDiagnoser]
public class CsvParallelDataReaderBenchmarks
{
    private string _path = null!;
    private CsvSchema _officeSchema = null!;
    private Dictionary<string, Type> _dataplatColumnTypes = null!;
    private CsvReaderChecksum _expected;

    [Params(100_000)]
    public int RowCount { get; set; }

    [Params(4)]
    public int MaxDegreeOfParallelism { get; set; }

    [Params(4096)]
    public int BatchSize { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        _path = Path.Combine(Path.GetTempPath(), $"officeimo-csv-parallel-reader-{Guid.NewGuid():N}.csv");
        long idSum = 0;
        decimal amountSum = 0;
        int trueCount = 0;
        long dateTicks = 0;
        long nameLength = 0;
        ulong orderedIdHash = 14695981039346656037UL;
        using (var writer = new StreamWriter(
                   _path,
                   append: false,
                   new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
                   bufferSize: 128 * 1024))
        {
            writer.WriteLine("Id,Amount,Active,Created,Name");
            for (int id = 1; id <= RowCount; id++)
            {
                int day = (id % 28) + 1;
                decimal amount = id * 1.25m;
                bool active = (id & 1) == 0;
                string name = $"Row {id.ToString(CultureInfo.InvariantCulture)}";
                writer.Write(id.ToString(CultureInfo.InvariantCulture));
                writer.Write(',');
                writer.Write(amount.ToString(CultureInfo.InvariantCulture));
                writer.Write(',');
                writer.Write(active ? "true" : "false");
                writer.Write(',');
                writer.Write("2026-08-");
                writer.Write(day.ToString("00", CultureInfo.InvariantCulture));
                writer.Write(',');
                writer.WriteLine(name);

                idSum += id;
                amountSum += amount;
                if (active) trueCount++;
                dateTicks += new DateTime(2026, 8, day).Ticks;
                nameLength += name.Length;
                orderedIdHash = unchecked((orderedIdHash ^ (uint)id) * 1099511628211UL);
            }
        }

        _officeSchema = new CsvSchemaBuilder()
            .Column("Id").AsInt32()
            .Column("Amount").AsType(typeof(decimal))
            .Column("Active").AsBoolean()
            .Column("Created").AsDateTime()
            .Column("Name").AsString()
            .Done()
            .Build();
        _dataplatColumnTypes = new Dictionary<string, Type>(StringComparer.OrdinalIgnoreCase)
        {
            ["Id"] = typeof(int),
            ["Amount"] = typeof(decimal),
            ["Active"] = typeof(bool),
            ["Created"] = typeof(DateTime),
            ["Name"] = typeof(string)
        };

        _expected = new CsvReaderChecksum(
            RowCount,
            idSum,
            amountSum,
            trueCount,
            dateTicks,
            nameLength,
            orderedIdHash);
        Validate(nameof(OfficeIMOSequential), OfficeIMOSequential());
        Validate(nameof(OfficeIMOParallel), OfficeIMOParallel());
        Validate(nameof(DataplatSequential), DataplatSequential());
        Validate(nameof(DataplatParallel), DataplatParallel());
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        if (File.Exists(_path)) File.Delete(_path);
    }

    [Benchmark(Description = "OfficeIMO-Sequential")]
    public CsvReaderChecksum OfficeIMOSequential()
    {
        using var reader = CsvDocument.OpenDataReader(
            _path,
            readerOptions: new CsvDataReaderOptions { Schema = _officeSchema });
        return Consume(reader);
    }

    [Benchmark(Description = "OfficeIMO-Parallel")]
    public CsvReaderChecksum OfficeIMOParallel()
    {
        using var reader = CsvDocument.OpenDataReader(
            _path,
            readerOptions: new CsvDataReaderOptions
            {
                Schema = _officeSchema,
                ParallelProcessing = new CsvDataReaderParallelOptions
                {
                    MaxDegreeOfParallelism = MaxDegreeOfParallelism,
                    BatchSize = BatchSize
                }
            });
        return Consume(reader);
    }

    [Benchmark(Description = "Dataplat-Sequential")]
    public CsvReaderChecksum DataplatSequential()
    {
        using var reader = new DataplatCsvDataReader(_path, CreateDataplatOptions(enableParallel: false));
        return Consume(reader);
    }

    [Benchmark(Baseline = true, Description = "Dataplat-Parallel")]
    public CsvReaderChecksum DataplatParallel()
    {
        using var reader = new DataplatCsvDataReader(_path, CreateDataplatOptions(enableParallel: true));
        return Consume(reader);
    }

    private DataplatCsvReaderOptions CreateDataplatOptions(bool enableParallel) => new()
    {
        HasHeaderRow = true,
        Culture = CultureInfo.InvariantCulture,
        ColumnTypes = _dataplatColumnTypes,
        EnableParallelProcessing = enableParallel,
        MaxDegreeOfParallelism = MaxDegreeOfParallelism,
        ParallelBatchSize = BatchSize
    };

    private static CsvReaderChecksum Consume(System.Data.IDataReader reader)
    {
        long idSum = 0;
        decimal amountSum = 0;
        int trueCount = 0;
        long dateTicks = 0;
        long nameLength = 0;
        ulong orderedIdHash = 14695981039346656037UL;
        int rows = 0;
        var values = new object[reader.FieldCount];
        while (reader.Read())
        {
            if (reader.GetValues(values) != values.Length)
            {
                throw new InvalidOperationException("The reader did not return every projected field.");
            }

            rows++;
            int id = (int)values[0];
            idSum += id;
            orderedIdHash = unchecked((orderedIdHash ^ (uint)id) * 1099511628211UL);
            amountSum += (decimal)values[1];
            if ((bool)values[2]) trueCount++;
            dateTicks += ((DateTime)values[3]).Ticks;
            nameLength += ((string)values[4]).Length;
        }

        return new CsvReaderChecksum(rows, idSum, amountSum, trueCount, dateTicks, nameLength, orderedIdHash);
    }

    private void Validate(string method, CsvReaderChecksum actual)
    {
        if (actual != _expected)
        {
            throw new InvalidOperationException(
                $"{method} returned {actual}; expected {_expected}.");
        }
    }

    public readonly record struct CsvReaderChecksum(
        int Rows,
        long IdSum,
        decimal AmountSum,
        int TrueCount,
        long DateTicks,
        long NameLength,
        ulong OrderedIdHash);
}
