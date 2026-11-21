using System.Text;
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
        TomlWorkloadCheck.RecordParse(Workload);
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:Tomlyn", "case:Default")]
    public long Tomlyn() => Sum(TomlLibraries.ParseTomlyn);

    [Benchmark]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
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
    private byte[][] _utf8Inputs = Array.Empty<byte[]>();

    [ParamsAllValues]
    public TomlWorkload Workload { get; set; }

    [GlobalSetup]
    public void Setup() {
        _inputs = TomlWorkloads.Get(Workload);
        _utf8Inputs = _inputs.Select(Encoding.UTF8.GetBytes).ToArray();
        TomlWorkloadCheck.Validate(Workload);
        TomlWorkloadCheck.RecordLoad(Workload);
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:Tomlyn", "case:Default")]
    public int Tomlyn() {
        int total = 0;
        foreach (string input in _inputs) total += TomlLibraries.LoadTomlyn(input).Count;
        return total;
    }

    [Benchmark]
    [BenchmarkCategory("library:Tomlet", "case:Default")]
    public int Tomlet() {
        int total = 0;
        foreach (string input in _inputs) total += TomlLibraries.LoadTomlet(input).Entries.Count;
        return total;
    }

    /// <summary>CsToml reads UTF-8 bytes, its native input, encoded once during setup.</summary>
    [Benchmark]
    [BenchmarkCategory("library:CsToml", "case:Default")]
    public int CsToml() {
        int total = 0;
        foreach (byte[] input in _utf8Inputs) total += TomlLibraries.LoadCsToml(input).RootNode.HasValue ? 1 : 0;
        return total;
    }

    /// <summary>OfficeIMO decodes values on first read; this lane leaves them undecoded.</summary>
    [Benchmark]
    [BenchmarkCategory("library:OfficeIMO", "case:LazyValues")]
    public int OfficeIMOToml() {
        int total = 0;
        foreach (string input in _inputs) total += TomlLibraries.LoadOfficeIMO(input).Root.Count;
        return total;
    }

    /// <summary>Also decodes every value, which the reference loaders do while loading.</summary>
    [Benchmark]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public long OfficeIMOTomlWithValues() {
        long total = 0;
        foreach (string input in _inputs) total += TomlLibraries.CountOfficeIMO(TomlLibraries.LoadOfficeIMO(input).Root);
        return total;
    }
}

/// <summary>
/// Load and save without edits. OfficeIMO and Tomlyn's syntax tree reproduce the input, which setup checks. Tomlet only
/// re-serializes its model, which is not the same work, so it has no lane here.
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
        TomlWorkloadCheck.RecordLoadAndSave(Workload);
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:Tomlyn", "case:Default")]
    public long TomlynSyntax() => Sum(TomlLibraries.SaveTomlyn);

    [Benchmark]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public long OfficeIMOToml() => Sum(TomlLibraries.SaveOfficeIMO);

    private long Sum(Func<string, string> save) {
        long total = 0;
        foreach (string input in _inputs) total += save(input).Length;
        return total;
    }
}
