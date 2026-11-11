using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Toml.Benchmarks;

/// <summary>One pass over the text without building a model.</summary>
[MemoryDiagnoser]
public class TomlParseBenchmarks {
    private string[] _inputs = Array.Empty<string>();

    [ParamsAllValues]
    public TomlWorkload Workload { get; set; }

    [GlobalSetup]
    public void Setup() {
        _inputs = TomlWorkloads.Get(Workload);
        TomlWorkloadCheck.Validate(Workload);
    }

    [Benchmark(Baseline = true)]
    public long Tomlyn() => Sum(TomlLibraries.ParseTomlyn);

    [Benchmark]
    public long OfficeIMOToml() => Sum(TomlLibraries.ParseOfficeIMO);

    private long Sum(Func<string, long> parse) {
        long total = 0;
        foreach (string input in _inputs) total += parse(input);
        return total;
    }
}

/// <summary>Load each library's document model.</summary>
[MemoryDiagnoser]
public class TomlLoadBenchmarks {
    private string[] _inputs = Array.Empty<string>();

    [ParamsAllValues]
    public TomlWorkload Workload { get; set; }

    [GlobalSetup]
    public void Setup() {
        _inputs = TomlWorkloads.Get(Workload);
        TomlWorkloadCheck.Validate(Workload);
    }

    [Benchmark(Baseline = true)]
    public int Tomlyn() {
        int total = 0;
        foreach (string input in _inputs) total += TomlLibraries.LoadTomlyn(input).Count;
        return total;
    }

    [Benchmark]
    public int Tomlet() {
        int total = 0;
        foreach (string input in _inputs) total += TomlLibraries.LoadTomlet(input).Entries.Count;
        return total;
    }

    /// <summary>OfficeIMO decodes values on first read; this lane leaves them undecoded.</summary>
    [Benchmark]
    public int OfficeIMOToml() {
        int total = 0;
        foreach (string input in _inputs) total += TomlLibraries.LoadOfficeIMO(input).Root.Count;
        return total;
    }

    /// <summary>Also decodes every value, which the reference loaders do while loading.</summary>
    [Benchmark]
    public long OfficeIMOTomlWithValues() {
        long total = 0;
        foreach (string input in _inputs) total += TomlLibraries.CountOfficeIMO(TomlLibraries.LoadOfficeIMO(input).Root);
        return total;
    }
}

/// <summary>
/// Load and save without edits. OfficeIMO and Tomlyn's syntax tree reproduce the input; Tomlet re-serializes it and
/// cannot write one corpus key, so its lane reports NA for the Corpus workload.
/// </summary>
[MemoryDiagnoser]
public class TomlLoadAndSaveBenchmarks {
    private string[] _inputs = Array.Empty<string>();

    [ParamsAllValues]
    public TomlWorkload Workload { get; set; }

    [GlobalSetup]
    public void Setup() {
        _inputs = TomlWorkloads.Get(Workload);
        TomlWorkloadCheck.Validate(Workload);
    }

    [Benchmark(Baseline = true)]
    public long TomlynSyntax() => Sum(TomlLibraries.SaveTomlyn);

    [Benchmark]
    public long TomletReserialized() => Sum(TomlLibraries.SaveTomlet);

    [Benchmark]
    public long OfficeIMOToml() => Sum(TomlLibraries.SaveOfficeIMO);

    private long Sum(Func<string, string> save) {
        long total = 0;
        foreach (string input in _inputs) total += save(input).Length;
        return total;
    }
}
