using System.Runtime.CompilerServices;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.Excel.Benchmarks;

internal static class BenchmarkProcessPriority {
    [ModuleInitializer]
    internal static void ApplyConfiguredPriority() {
        string? priority = Environment.GetEnvironmentVariable(
            "OFFICEIMO_BENCHMARK_PROCESS_PRIORITY");
        if (!string.IsNullOrEmpty(priority)) {
            BenchmarkProcessorAffinity.ApplyPriority(priority);
        }
    }
}
