using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>Measures first-row initialization against a million-row input using the validated async workload.</summary>
[MemoryDiagnoser]
public class CsvLargeFirstRowBenchmarks
{
    private CsvAsyncReadBenchmarks _workload = null!;

    [Params(CsvAsyncReadBenchmarks.AsyncReadShape.Plain, CsvAsyncReadBenchmarks.AsyncReadShape.Multiline)]
    public CsvAsyncReadBenchmarks.AsyncReadShape Shape { get; set; }

    [GlobalSetup]
    public Task Setup()
    {
        // BDN may raise worker priority; this qualification lane fixes it explicitly.
        BenchmarkProcessorAffinity.ApplyPriority("Normal");
        _workload = new CsvAsyncReadBenchmarks {
            RowCount = 1_000_000, Shape = Shape,
            Operation = CsvAsyncReadBenchmarks.AsyncReadOperation.FirstRow
        };
        return _workload.Setup();
    }

    [GlobalCleanup]
    public void Cleanup() => _workload.Cleanup();

    [Benchmark(Baseline = true)]
    public Task<CsvAsyncReadBenchmarks.ReadChecksum> OfficeIMO_Snapshot() => _workload.OfficeIMO_Snapshot();

    [Benchmark]
    public Task<CsvAsyncReadBenchmarks.ReadChecksum> OfficeIMO_Incremental() => _workload.OfficeIMO_Incremental();
}
