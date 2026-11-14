#if NETFRAMEWORK
using BenchmarkDotNet.Running;

namespace OfficeIMO.Word.Benchmarks;

/// <summary>
/// .NET Framework entry point. Windows PowerShell 5.1 runs PSWriteOffice on .NET Framework, so the create, read and
/// replace comparisons build for net472 too; the rest of the suite needs .NET 8 or later.
/// </summary>
internal static class NetFrameworkProgram {
    private static void Main(string[] args) =>
        BenchmarkSwitcher.FromTypes([
            typeof(WordCreateParagraphComparisonBenchmarks),
            typeof(WordCreateReportComparisonBenchmarks),
            typeof(WordReadComparisonBenchmarks),
            typeof(WordReplaceComparisonBenchmarks)
        ]).Run(args);
}
#endif
