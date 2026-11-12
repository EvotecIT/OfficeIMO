using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Yaml.Benchmarks.Kit;

/// <summary>
/// The oneyaml-kit tasks that leave their starting state unchanged (typed load, generic load, dump, document parse,
/// unchanged emit) on the kit's workloads, one benchmark per library column of <see cref="YamlKitLanes"/>. A library
/// without the task, or without that workload, fails its setup and shows no number. The edit script consumes its
/// document, so it runs in <c>--kit-timings</c> only. <c>OFFICEIMO_YAML_KIT_WORKLOADS</c> names the workloads; select
/// cases with <c>--filter</c> or <c>OFFICEIMO_YAML_KIT_CASES</c> (comma-separated <c>Task/workload</c>).
/// </summary>
[MemoryDiagnoser]
public class YamlKitBenchmarks {
    private YamlKitLane _lane = null!;
    private object _state = null!;

    [ParamsSource(nameof(Cases))]
    public string Case { get; set; } = string.Empty;

    public static IEnumerable<string> Cases() {
        string? selected = Environment.GetEnvironmentVariable("OFFICEIMO_YAML_KIT_CASES");
        if (selected is { Length: > 0 }) return selected.Split(',', StringSplitOptions.RemoveEmptyEntries);
        return from task in Enum.GetValues<YamlKitTask>()
               where task != YamlKitTask.EditScript
               from workload in YamlKitWorkload.All
               where workload.Tasks.Contains(task)
               select task + "/" + workload.Id;
    }

    [GlobalSetup(Target = nameof(OfficeIMO))] public void SetupOfficeIMO() => Setup("OfficeIMO");
    [GlobalSetup(Target = nameof(OfficeIMOReflection))] public void SetupOfficeIMOReflection() => Setup("OfficeIMO reflection");
    [GlobalSetup(Target = nameof(OfficeIMOUtf8))] public void SetupOfficeIMOUtf8() => Setup("OfficeIMO UTF-8");
    [GlobalSetup(Target = nameof(OneYaml))] public void SetupOneYaml() => Setup("OneYaml");
    [GlobalSetup(Target = nameof(OneYamlGenerated))] public void SetupOneYamlGenerated() => Setup("OneYaml generated");
    [GlobalSetup(Target = nameof(OneYamlString))] public void SetupOneYamlString() => Setup("OneYaml string");
    [GlobalSetup(Target = nameof(VYaml))] public void SetupVYaml() => Setup("VYaml");
    [GlobalSetup(Target = nameof(SharpYaml))] public void SetupSharpYaml() => Setup("SharpYaml");

    [Benchmark] public object? OfficeIMO() => _lane.Run(_state);
    [Benchmark] public object? OfficeIMOReflection() => _lane.Run(_state);
    [Benchmark] public object? OfficeIMOUtf8() => _lane.Run(_state);
    [Benchmark] public object? OneYaml() => _lane.Run(_state);
    [Benchmark] public object? OneYamlGenerated() => _lane.Run(_state);
    [Benchmark] public object? OneYamlString() => _lane.Run(_state);
    [Benchmark] public object? VYaml() => _lane.Run(_state);
    [Benchmark] public object? SharpYaml() => _lane.Run(_state);

    private void Setup(string library) {
        int slash = Case.IndexOf('/');
        YamlKitTask task = Enum.Parse<YamlKitTask>(Case.Substring(0, slash));
        YamlKitWorkload workload = YamlKitWorkload.Get(Case.Substring(slash + 1));
        _lane = YamlKitLanes.Find(library, task, workload) ?? throw new NotSupportedException(library + " has no " + task + " lane for " + workload.Id + ".");
        _lane.Validate(workload);
        _state = _lane.Prepare(workload);
    }
}
