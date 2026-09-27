using System.Data.Common;
using System.Globalization;
using System.Text;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;
using System.Diagnostics;

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>
/// Measures the public snapshot and incremental APIs for the same first-row or full-traversal
/// result. Snapshot construction intentionally reads the whole input even in the first-row case.
/// Setup validates every returned field against generated values before timing.
/// </summary>
[MemoryDiagnoser]
public class CsvAsyncReadBenchmarks
{
    private string _path = null!;
    private ReadChecksum _expected;

    [Params(1_000, 100_000)]
    public int RowCount { get; set; }

    [Params(AsyncReadShape.Plain, AsyncReadShape.Multiline)]
    public AsyncReadShape Shape { get; set; }

    [Params(AsyncReadOperation.FirstRow, AsyncReadOperation.AllRows)]
    public AsyncReadOperation Operation { get; set; }

    [GlobalSetup]
    public async Task Setup()
    {
        using (var process = Process.GetCurrentProcess())
        {
            string? expectedAffinity = Environment.GetEnvironmentVariable("OFFICEIMO_EXPECTED_BENCHMARK_AFFINITY");
            if (!string.IsNullOrEmpty(expectedAffinity))
            {
                if (!OperatingSystem.IsWindows() && !OperatingSystem.IsLinux())
                    throw new PlatformNotSupportedException("Processor affinity qualification requires Windows or Linux.");
                if (process.ProcessorAffinity != BenchmarkProcessorAffinity.ParseList(expectedAffinity)[0])
                    throw new InvalidOperationException("The benchmark worker did not inherit its declared processor affinity.");
            }
            Console.WriteLine(OperatingSystem.IsWindows()
                ? $"Worker placement: affinity={BenchmarkProcessorAffinity.Format(process.ProcessorAffinity)}; priority={process.PriorityClass}"
                : $"Worker placement: priority={process.PriorityClass}");
        }
        _path = Path.Combine(Path.GetTempPath(), $"officeimo-csv-async-read-{Guid.NewGuid():N}.csv");
        using (var writer = new StreamWriter(_path, false, new UTF8Encoding(false)))
        {
            writer.WriteLine("Id,Name,Notes");
            for (int id = 1; id <= RowCount; id++)
            {
                var values = Values(id);
                writer.Write(values.Id);
                writer.Write(',');
                writer.Write(values.Name);
                writer.Write(',');
                writer.Write('"');
                writer.Write(values.Notes.Replace("\"", "\"\""));
                writer.WriteLine('"');
            }
        }
        int observed = Operation == AsyncReadOperation.FirstRow ? 1 : RowCount;
        long idSum = 0, characters = 0;
        for (int id = 1; id <= observed; id++)
        {
            var values = Values(id);
            idSum += id;
            characters += values.Id.Length + values.Name.Length + values.Notes.Length;
        }
        _expected = new ReadChecksum(observed, idSum, characters);
        await ValidateAsync(incremental: false);
        await ValidateAsync(incremental: true);
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        if (File.Exists(_path)) File.Delete(_path);
    }

    [Benchmark(Baseline = true)]
    public async Task<ReadChecksum> OfficeIMO_Snapshot()
    {
        using var reader = await CsvDocument.OpenDataReaderAsync(_path).ConfigureAwait(false);
        return await ConsumeAsync(reader).ConfigureAwait(false);
    }

    [Benchmark]
    public async Task<ReadChecksum> OfficeIMO_Incremental()
    {
        using var reader = await CsvDocument.OpenStreamingDataReaderAsync(_path).ConfigureAwait(false);
        return await ConsumeAsync(reader).ConfigureAwait(false);
    }

    private async Task<ReadChecksum> ConsumeAsync(DbDataReader reader)
    {
        int rows = 0;
        long idSum = 0, characters = 0;
        while (await reader.ReadAsync().ConfigureAwait(false))
        {
            rows++;
            string id = reader.GetString(0);
            idSum += int.Parse(id, CultureInfo.InvariantCulture);
            characters += id.Length + reader.GetString(1).Length + reader.GetString(2).Length;
            if (Operation == AsyncReadOperation.FirstRow) break;
        }
        return new ReadChecksum(rows, idSum, characters);
    }

    private async Task ValidateAsync(bool incremental)
    {
        using var reader = incremental
            ? await CsvDocument.OpenStreamingDataReaderAsync(_path).ConfigureAwait(false)
            : await CsvDocument.OpenDataReaderAsync(_path).ConfigureAwait(false);
        int rows = 0;
        while (await reader.ReadAsync().ConfigureAwait(false))
        {
            var expected = Values(++rows);
            if (reader.GetString(0) != expected.Id || reader.GetString(1) != expected.Name || reader.GetString(2) != expected.Notes)
                throw new InvalidDataException($"CSV async {incremental} field mismatch at row {rows}.");
            if (Operation == AsyncReadOperation.FirstRow) break;
        }
        ReadChecksum actual = incremental
            ? await OfficeIMO_Incremental().ConfigureAwait(false)
            : await OfficeIMO_Snapshot().ConfigureAwait(false);
        if (rows != _expected.Rows || actual != _expected)
            throw new InvalidDataException($"CSV async {incremental} checksum mismatch: {actual} versus {_expected}.");
    }

    private (string Id, string Name, string Notes) Values(int id)
    {
        string key = id.ToString(CultureInfo.InvariantCulture);
        return (key, "Person " + key, Shape == AsyncReadShape.Plain
            ? "Unique note " + key
            : "Row " + key + "\r\nsecond line, \"quoted\" Zażółć");
    }

    public enum AsyncReadShape { Plain, Multiline }
    public enum AsyncReadOperation { FirstRow, AllRows }
    public readonly record struct ReadChecksum(int Rows, long IdSum, long Characters);
}
