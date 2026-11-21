using BenchmarkDotNet.Filters;
using BenchmarkDotNet.Running;

namespace OfficeIMO.Yaml.Benchmarks.Workloads;

/// <summary>
/// Leaves out the SharpYaml save lanes whose output no other reader accepts. SharpYaml writes %TAG directives before
/// each later document without the "..." end marker YAML requires there, so its multi-document output for these
/// workloads fails setup validation; Build/LibraryComparisonDefinitions.ps1 leaves the same cases out.
/// It also leaves out VYaml's ManySmallFiles load: VYaml resolves the plain scalar 'infinity' in one corpus file as a
/// float, so its loaded values differ from the other loaders and the comparison declares no cell for it.
/// </summary>
internal static class YamlSaveLaneFilter {
    private static readonly HashSet<string> SaveBenchmarks = new(StringComparer.Ordinal) {
        nameof(YamlLoadAndSaveBenchmarks),
        nameof(YamlEditAndSaveBenchmarks),
        nameof(YamlEmitBenchmarks)
    };

    private static readonly HashSet<YamlWorkload> MultiDocumentWorkloads = new() {
        YamlWorkload.KubernetesStream,
        YamlWorkload.DeepNesting,
        YamlWorkload.ManySmallFiles
    };

    internal static IFilter Instance { get; } = new SimpleFilter(ShouldRun);

    private static bool ShouldRun(BenchmarkCase benchmarkCase) {
        if (benchmarkCase.Descriptor.Type.Name == nameof(YamlLoadTreeBenchmarks) &&
            benchmarkCase.Descriptor.WorkloadMethod.Name == nameof(YamlLoadTreeBenchmarks.VYaml)) {
            return !(benchmarkCase.Parameters["Workload"] is YamlWorkload.ManySmallFiles);
        }

        if (!SaveBenchmarks.Contains(benchmarkCase.Descriptor.Type.Name) ||
            benchmarkCase.Descriptor.WorkloadMethod.Name != nameof(YamlLoadAndSaveBenchmarks.SharpYaml)) {
            return true;
        }

        return !(benchmarkCase.Parameters["Workload"] is YamlWorkload workload && MultiDocumentWorkloads.Contains(workload));
    }
}
