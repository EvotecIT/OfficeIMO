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

    [Benchmark(Baseline = true)]
    public long YamlDotNet() => Run(YamlReferenceLibraries.YamlDotNet);

    [Benchmark]
    public long SharpYaml() => Run(YamlReferenceLibraries.SharpYaml);

    [Benchmark]
    public long OfficeIMOYaml() => Run(YamlReferenceLibraries.OfficeIMOYaml);

    private void Setup(IYamlReferenceLibrary library) {
        _inputs = YamlWorkloads.Get(Workload);
        YamlWorkloadValidation.ValidateParse(library, Workload, _inputs);
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

    [Benchmark(Baseline = true)]
    public long YamlDotNet() => Run(YamlReferenceLibraries.YamlDotNet);

    [Benchmark]
    public long SharpYaml() => Run(YamlReferenceLibraries.SharpYaml);

    private void Setup(IYamlReferenceLibrary library) {
        _inputs = YamlWorkloads.Get(Workload);
        YamlWorkloadValidation.ValidateLoad(library, Workload, _inputs);
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

    [Benchmark(Baseline = true)]
    public long YamlDotNet() => Run(YamlReferenceLibraries.YamlDotNet);

    [Benchmark]
    public long SharpYaml() => Run(YamlReferenceLibraries.SharpYaml);

    private void Setup(IYamlReferenceLibrary library) {
        _inputs = YamlWorkloads.Get(Workload);
        YamlWorkloadValidation.ValidateLoadAndSave(library, Workload, _inputs);
    }

    private long Run(IYamlReferenceLibrary library) {
        long total = 0;
        foreach (string input in _inputs) total += library.LoadAndSave(input).Length;
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

    [Benchmark(Baseline = true)]
    public long YamlDotNet() => Run(YamlReferenceLibraries.YamlDotNet);

    [Benchmark]
    public long SharpYaml() => Run(YamlReferenceLibraries.SharpYaml);

    private void Setup(IYamlReferenceLibrary library) {
        string[] inputs = YamlWorkloads.Get(Workload);
        _events = inputs.Select(library.CaptureEvents).ToArray();
        YamlWorkloadValidation.ValidateEmit(library, Workload, inputs, _events);
    }

    private long Run(IYamlReferenceLibrary library) {
        long total = 0;
        foreach (object events in _events) total += library.Emit(events).Length;
        return total;
    }
}
