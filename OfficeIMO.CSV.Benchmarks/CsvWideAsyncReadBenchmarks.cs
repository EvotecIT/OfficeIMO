using System.Data.Common;
using System.Diagnostics;
using System.Text;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>
/// Measures the public snapshot and incremental readers on wide rows whose values are distinct
/// across every row and column. Fixture generation and full field validation are outside timing.
/// </summary>
[MemoryDiagnoser]
public class CsvWideAsyncReadBenchmarks
{
    private const int FieldCount = 32;
    private string _path = null!;
    private ReadChecksum _expected;

    [Params(5_000, 25_000)]
    public int RowCount { get; set; }

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
                if (OperatingSystem.IsWindows() || OperatingSystem.IsLinux())
                {
                    if (process.ProcessorAffinity != BenchmarkProcessorAffinity.ParseList(expectedAffinity)[0])
                        throw new InvalidOperationException("The benchmark worker did not inherit its declared processor affinity.");
                }
                else
                    throw new PlatformNotSupportedException("Processor affinity qualification requires Windows or Linux.");
            }
            Console.WriteLine(OperatingSystem.IsWindows()
                ? $"Worker placement: affinity={BenchmarkProcessorAffinity.Format(process.ProcessorAffinity)}; priority={process.PriorityClass}"
                : $"Worker placement: priority={process.PriorityClass}");
        }

        _path = Path.Combine(Path.GetTempPath(), $"officeimo-csv-wide-read-{Guid.NewGuid():N}.csv");
        await using (var stream = new FileStream(_path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
        await using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
        {
            for (int column = 0; column < FieldCount; column++)
            {
                if (column != 0) await writer.WriteAsync(',').ConfigureAwait(false);
                await writer.WriteAsync($"Field{column:D2}").ConfigureAwait(false);
            }
            await writer.WriteLineAsync().ConfigureAwait(false);
            for (int row = 0; row < RowCount; row++)
            {
                for (int column = 0; column < FieldCount; column++)
                {
                    if (column != 0) await writer.WriteAsync(',').ConfigureAwait(false);
                    await writer.WriteAsync(Value(row, column)).ConfigureAwait(false);
                }
                await writer.WriteLineAsync().ConfigureAwait(false);
            }
        }

        int observedRows = Operation == AsyncReadOperation.FirstRow ? 1 : RowCount;
        long characters = 0;
        long signature = 0;
        for (int row = 0; row < observedRows; row++)
        for (int column = 0; column < FieldCount; column++)
        {
            string value = Value(row, column);
            characters += value.Length;
            signature += value[0] + value[^1];
        }
        _expected = new ReadChecksum(observedRows, observedRows * FieldCount, characters, signature);
        await ValidateAsync(incremental: false).ConfigureAwait(false);
        await ValidateAsync(incremental: true).ConfigureAwait(false);
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        if (File.Exists(_path)) File.Delete(_path);
    }

    [Benchmark(Baseline = true)]
    public async Task<ReadChecksum> OfficeIMO_Snapshot()
    {
        var document = await CsvDocument.LoadAsync(_path).ConfigureAwait(false);
        using var reader = document.CreateDataReader();
        return await ConsumeAsync(reader).ConfigureAwait(false);
    }

    [Benchmark]
    public async Task<ReadChecksum> OfficeIMO_Incremental()
    {
        using var reader = await CsvDocument.OpenDataReaderAsync(_path).ConfigureAwait(false);
        return await ConsumeAsync(reader).ConfigureAwait(false);
    }

    private async Task<ReadChecksum> ConsumeAsync(DbDataReader reader)
    {
        int rows = 0;
        long characters = 0;
        long signature = 0;
        while (await reader.ReadAsync().ConfigureAwait(false))
        {
            rows++;
            for (int column = 0; column < FieldCount; column++)
            {
                string value = reader.GetString(column);
                characters += value.Length;
                signature += value[0] + value[^1];
            }
            if (Operation == AsyncReadOperation.FirstRow) break;
        }
        return new ReadChecksum(rows, rows * FieldCount, characters, signature);
    }

    private async Task ValidateAsync(bool incremental)
    {
        using var reader = incremental
            ? await CsvDocument.OpenDataReaderAsync(_path).ConfigureAwait(false)
            : (await CsvDocument.LoadAsync(_path).ConfigureAwait(false)).CreateDataReader();
        int rows = 0;
        while (await reader.ReadAsync().ConfigureAwait(false))
        {
            for (int column = 0; column < FieldCount; column++)
                if (reader.GetString(column) != Value(rows, column))
                    throw new InvalidDataException($"CSV wide {incremental} field mismatch at row {rows}, column {column}.");
            rows++;
            if (Operation == AsyncReadOperation.FirstRow) break;
        }
        ReadChecksum actual = incremental
            ? await OfficeIMO_Incremental().ConfigureAwait(false)
            : await OfficeIMO_Snapshot().ConfigureAwait(false);
        if (rows != _expected.Rows || actual != _expected)
            throw new InvalidDataException($"CSV wide {incremental} checksum mismatch: {actual} versus {_expected}.");
    }

    private static string Value(int row, int column) => $"v{row:D6}_{column:D2}";

    public enum AsyncReadOperation { FirstRow, AllRows }
    public readonly record struct ReadChecksum(int Rows, int Cells, long Characters, long Signature);
}
