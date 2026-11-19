using BenchmarkDotNet.Attributes;
using OfficeIMO.Yaml.Benchmarks.Workloads;

namespace OfficeIMO.Yaml.Benchmarks;

/// <summary>Parse a stream to events without building a tree.</summary>
[MemoryDiagnoser]
public class YamlParseEventsBenchmarks {
    private string[] _inputs = Array.Empty<string>();

    [ParamsAllValues]
    public YamlWorkload Workload { get; set; }

    [GlobalSetup(Target = nameof(YamlDotNet))]
    public void SetupYamlDotNet() => Setup(YamlReferenceLibraries.YamlDotNet);

    [GlobalSetup(Target = nameof(SharpYaml))]
    public void SetupSharpYaml() => Setup(YamlReferenceLibraries.SharpYaml);

    [GlobalSetup(Target = nameof(OfficeIMOYaml))]
    public void SetupOfficeIMOYaml() => Setup(YamlReferenceLibraries.OfficeIMOYaml);

    [GlobalSetup(Target = nameof(VYaml))]
    public void SetupVYaml() => Setup(VYamlLibrary.Instance);

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:YamlDotNet", "case:Default")]
    public long YamlDotNet() => Run(YamlReferenceLibraries.YamlDotNet);

    [Benchmark]
    [BenchmarkCategory("library:SharpYaml", "case:Default")]
    public long SharpYaml() => Run(YamlReferenceLibraries.SharpYaml);

    [Benchmark]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public long OfficeIMOYaml() => Run(YamlReferenceLibraries.OfficeIMOYaml);

    /// <summary>VYaml reads UTF-8 bytes encoded once during setup.</summary>
    [Benchmark]
    [BenchmarkCategory("library:VYaml", "case:Default")]
    public long VYaml() => Run(VYamlLibrary.Instance);

    private void Setup(IYamlReferenceLibrary library) {
        _inputs = YamlWorkloads.Get(Workload);
        YamlWorkloadValidation.Record(library, Workload, "event count matches the library's conformance adapter; event content matches {other}",
            () => YamlWorkloadValidation.ValidateParse(library, Workload, _inputs));
    }

    private long Run(IYamlReferenceLibrary library) {
        long total = 0;
        foreach (string input in _inputs) total += library.ParseEvents(input);
        return total;
    }
}

/// <summary>Load a stream into the library's representation tree.</summary>
[MemoryDiagnoser]
public class YamlLoadTreeBenchmarks {
    private string[] _inputs = Array.Empty<string>();

    [ParamsAllValues]
    public YamlWorkload Workload { get; set; }

    [GlobalSetup(Target = nameof(YamlDotNet))]
    public void SetupYamlDotNet() => Setup(YamlReferenceLibraries.YamlDotNet);

    [GlobalSetup(Target = nameof(SharpYaml))]
    public void SetupSharpYaml() => Setup(YamlReferenceLibraries.SharpYaml);

    [GlobalSetup(Targets = new[] { nameof(OfficeIMOYaml), nameof(OfficeIMOYamlWithValues) })]
    public void SetupOfficeIMOYaml() => Setup(YamlReferenceLibraries.OfficeIMOYaml);

    [GlobalSetup(Target = nameof(VYaml))]
    public void SetupVYaml() => Setup(VYamlLibrary.Instance);

    [GlobalSetup(Target = nameof(OneYaml))]
    public void SetupOneYaml() => Setup(OneYamlLibrary.Instance);

    /// <summary>OneYaml's tree is its lossless document; this is its untyped deserialization from UTF-8 bytes.</summary>
    [Benchmark]
    [BenchmarkCategory("library:OneYaml", "case:Default")]
    public long OneYaml() => Run(OneYamlLibrary.Instance);

    /// <summary>VYaml has no representation tree; this is its dynamic deserialization into dictionaries and lists.</summary>
    [Benchmark]
    [BenchmarkCategory("library:VYaml", "case:Default")]
    public long VYaml() => Run(VYamlLibrary.Instance);

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:YamlDotNet", "case:Default")]
    public long YamlDotNet() => Run(YamlReferenceLibraries.YamlDotNet);

    [Benchmark]
    [BenchmarkCategory("library:SharpYaml", "case:Default")]
    public long SharpYaml() => Run(YamlReferenceLibraries.SharpYaml);

    [Benchmark]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public long OfficeIMOYaml() => Run(YamlReferenceLibraries.OfficeIMOYaml);

    /// <summary>Also reads every scalar value, which the reference trees materialize while loading.</summary>
    [Benchmark]
    [BenchmarkCategory("library:OfficeIMO", "case:WithValues")]
    public long OfficeIMOYamlWithValues() {
        long total = 0;
        foreach (string input in _inputs) total += YamlReferenceLibraries.OfficeIMOYaml.LoadTreeAndValues(input);
        return total;
    }

    private void Setup(IYamlReferenceLibrary library) {
        _inputs = YamlWorkloads.Get(Workload);
        // OfficeIMO's tree reads scalar values on demand; only its WithValues lane does the reference trees' work, so
        // only that case is recorded as validated.
        YamlWorkloadValidation.Record(library, Workload, "tree covers every node event; loaded values match {other}",
            () => YamlWorkloadValidation.ValidateLoad(library, Workload, _inputs),
            ReferenceEquals(library, YamlReferenceLibraries.OfficeIMOYaml) ? "WithValues" : "");
    }

    private long Run(IYamlReferenceLibrary library) {
        long total = 0;
        foreach (string input in _inputs) total += library.LoadTree(input);
        return total;
    }
}

/// <summary>Load a stream and save it again without edits.</summary>
[MemoryDiagnoser]
public class YamlLoadAndSaveBenchmarks {
    private string[] _inputs = Array.Empty<string>();

    [ParamsAllValues]
    public YamlWorkload Workload { get; set; }

    [GlobalSetup(Target = nameof(YamlDotNet))]
    public void SetupYamlDotNet() => Setup(YamlReferenceLibraries.YamlDotNet);

    [GlobalSetup(Target = nameof(SharpYaml))]
    public void SetupSharpYaml() => Setup(YamlReferenceLibraries.SharpYaml);

    [GlobalSetup(Target = nameof(OfficeIMOYaml))]
    public void SetupOfficeIMOYaml() => Setup(YamlReferenceLibraries.OfficeIMOYaml);

    [GlobalSetup(Target = nameof(OneYaml))]
    public void SetupOneYaml() => Setup(OneYamlLibrary.Instance);

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:YamlDotNet", "case:Default")]
    public long YamlDotNet() => Run(YamlReferenceLibraries.YamlDotNet);

    [Benchmark]
    [BenchmarkCategory("library:SharpYaml", "case:Default")]
    public long SharpYaml() => Run(YamlReferenceLibraries.SharpYaml);

    [Benchmark]
    [BenchmarkCategory("library:OneYaml", "case:Default")]
    public long OneYaml() => Run(OneYamlLibrary.Instance);

    [Benchmark]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public long OfficeIMOYaml() => Run(YamlReferenceLibraries.OfficeIMOYaml);

    private void Setup(IYamlReferenceLibrary library) {
        _inputs = YamlWorkloads.Get(Workload);
        YamlWorkloadValidation.Record(library, Workload, "saved output loads to the input's values through {other}; lossless savers return the input unchanged",
            () => YamlWorkloadValidation.ValidateLoadAndSave(library, Workload, _inputs));
    }

    private long Run(IYamlReferenceLibrary library) {
        long total = 0;
        foreach (string input in _inputs) total += library.LoadAndSave(input).Length;
        return total;
    }
}

/// <summary>Load a stream, change one value and save it.</summary>
[MemoryDiagnoser]
public class YamlEditAndSaveBenchmarks {
    private string[] _inputs = Array.Empty<string>();

    [ParamsAllValues]
    public YamlWorkload Workload { get; set; }

    [GlobalSetup(Target = nameof(YamlDotNet))]
    public void SetupYamlDotNet() => Setup(YamlReferenceLibraries.YamlDotNet);

    [GlobalSetup(Target = nameof(SharpYaml))]
    public void SetupSharpYaml() => Setup(YamlReferenceLibraries.SharpYaml);

    [GlobalSetup(Target = nameof(OfficeIMOYaml))]
    public void SetupOfficeIMOYaml() => Setup(YamlReferenceLibraries.OfficeIMOYaml);

    [GlobalSetup(Target = nameof(OneYaml))]
    public void SetupOneYaml() => Setup(OneYamlLibrary.Instance);

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:YamlDotNet", "case:Default")]
    public long YamlDotNet() => Run(YamlReferenceLibraries.YamlDotNet);

    [Benchmark]
    [BenchmarkCategory("library:SharpYaml", "case:Default")]
    public long SharpYaml() => Run(YamlReferenceLibraries.SharpYaml);

    [Benchmark]
    [BenchmarkCategory("library:OneYaml", "case:Default")]
    public long OneYaml() => Run(OneYamlLibrary.Instance);

    [Benchmark]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public long OfficeIMOYaml() => Run(YamlReferenceLibraries.OfficeIMOYaml);

    private void Setup(IYamlReferenceLibrary library) {
        _inputs = YamlWorkloads.Get(Workload);
        YamlWorkloadValidation.Record(library, Workload, "edited output loads through {other} to the input's values with the edited scalar changed",
            () => YamlWorkloadValidation.ValidateEditAndSave(library, Workload, _inputs));
    }

    private long Run(IYamlReferenceLibrary library) {
        long total = 0;
        foreach (string input in _inputs) total += library.EditAndSave(input).Length;
        return total;
    }
}

/// <summary>Emit a new document from events captured outside timing.</summary>
[MemoryDiagnoser]
public class YamlEmitBenchmarks {
    private object[] _events = Array.Empty<object>();

    [ParamsAllValues]
    public YamlWorkload Workload { get; set; }

    [GlobalSetup(Target = nameof(YamlDotNet))]
    public void SetupYamlDotNet() => Setup(YamlReferenceLibraries.YamlDotNet);

    [GlobalSetup(Target = nameof(SharpYaml))]
    public void SetupSharpYaml() => Setup(YamlReferenceLibraries.SharpYaml);

    [GlobalSetup(Target = nameof(OfficeIMOYaml))]
    public void SetupOfficeIMOYaml() => Setup(YamlReferenceLibraries.OfficeIMOYaml);

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:YamlDotNet", "case:Default")]
    public long YamlDotNet() => Run(YamlReferenceLibraries.YamlDotNet);

    [Benchmark]
    [BenchmarkCategory("library:SharpYaml", "case:Default")]
    public long SharpYaml() => Run(YamlReferenceLibraries.SharpYaml);

    [Benchmark]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public long OfficeIMOYaml() => Run(YamlReferenceLibraries.OfficeIMOYaml);

    private void Setup(IYamlReferenceLibrary library) {
        string[] inputs = YamlWorkloads.Get(Workload);
        _events = inputs.Select(library.CaptureEvents).ToArray();
        YamlWorkloadValidation.Record(library, Workload, "emitted output parses through {other} to the input's event content",
            () => YamlWorkloadValidation.ValidateEmit(library, Workload, inputs, _events));
    }

    private long Run(IYamlReferenceLibrary library) {
        long total = 0;
        foreach (object events in _events) total += library.Emit(events).Length;
        return total;
    }
}
