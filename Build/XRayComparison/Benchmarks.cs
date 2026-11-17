using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Engines;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Reports;

namespace OfficeIMO.XRayComparison;

/// <summary>
/// Warm-up 3 and 12 measured iterations; median, spread and allocation per operation. Benchmarks run in
/// process: the out-of-process toolchain rebuilds the whole OfficeIMO project graph without shared
/// compilation and exceeds its build timeout.
/// </summary>
internal sealed class ComparisonConfig : ManualConfig {
    public ComparisonConfig() {
        AddJob(Job.Default
            .WithWarmupCount(3)
            .WithIterationCount(12)
            .WithLaunchCount(1)
            .WithStrategy(RunStrategy.Throughput)
            .WithToolchain(BenchmarkDotNet.Toolchains.InProcess.Emit.InProcessEmitToolchain.Instance)
            .WithId("XRayComparison"));
        AddColumn(StatisticColumn.Median, StatisticColumn.Min, StatisticColumn.Max, StatisticColumn.StdDev);
        AddDiagnoser(BenchmarkDotNet.Diagnosers.MemoryDiagnoser.Default);
        WithOrderer(new BenchmarkDotNet.Order.DefaultOrderer(BenchmarkDotNet.Order.SummaryOrderPolicy.Declared));
        WithSummaryStyle(SummaryStyle.Default.WithRatioStyle(RatioStyle.Trend));
        WithOptions(ConfigOptions.JoinSummary);
    }
}

public abstract class ExtractionBenchmarksBase {
    private byte[] _bytes = [];
    private BenchCase _case = null!;

    [ParamsSource(nameof(Cases))]
    public string Case { get; set; } = string.Empty;

    public IEnumerable<string> Cases() => Corpus.Extraction.Select(item => item.Id);

    [GlobalSetup]
    public void Setup() {
        _case = Corpus.Get(Case);
        _bytes = File.ReadAllBytes(_case.Path);
    }

    private protected ExtractionOutput RunOfficeImo(bool markdown) => markdown ? Workloads.OfficeImoMarkdown(_bytes, _case) : Workloads.OfficeImoText(_bytes, _case);
    private protected ExtractionOutput RunXRay(bool markdown) => markdown ? Workloads.XRayMarkdown(_bytes, _case) : Workloads.XRayText(_bytes, _case);
}

[Config(typeof(ComparisonConfig))]
public class TextExtractionBenchmarks : ExtractionBenchmarksBase {
    [Benchmark(Baseline = true)] public ExtractionOutput OfficeIMO() => RunOfficeImo(markdown: false);
    [Benchmark] public ExtractionOutput XRay() => RunXRay(markdown: false);
}

[Config(typeof(ComparisonConfig))]
public class MarkdownExtractionBenchmarks : ExtractionBenchmarksBase {
    [Benchmark(Baseline = true)] public ExtractionOutput OfficeIMO() => RunOfficeImo(markdown: true);
    [Benchmark] public ExtractionOutput XRay() => RunXRay(markdown: true);
}

[Config(typeof(ComparisonConfig))]
public class OfficeToPdfBenchmarks {
    private BenchCase _case = null!;

    [ParamsSource(nameof(Cases))]
    public string Case { get; set; } = string.Empty;

    public IEnumerable<string> Cases() => Corpus.OfficeToPdf.Select(item => item.Id);

    [GlobalSetup] public void Setup() => _case = Corpus.Get(Case);

    [Benchmark(Baseline = true)] public byte[] OfficeIMO() => Workloads.OfficeImoPdf(_case);
    [Benchmark] public byte[] XRay() => Workloads.XRayPdf(_case);
}
