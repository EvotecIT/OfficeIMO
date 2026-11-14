#if NETFRAMEWORK
using BenchmarkDotNet.Running;

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>
/// .NET Framework entry point. Windows PowerShell 5.1 runs PSWriteOffice and DbaClientX on .NET Framework, so the
/// lanes that mirror those modules' workloads build for net472 too; the rest of the suite needs .NET 8 or later.
/// </summary>
internal static class NetFrameworkProgram {
    private static void Main(string[] args) =>
        BenchmarkSwitcher.FromTypes([typeof(CsvDbatoolsExportBenchmarks)]).Run(args);
}
#endif
