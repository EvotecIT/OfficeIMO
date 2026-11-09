using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>Materializes and formats the same 50,000-by-15 logical values through loaded workbook APIs.</summary>
[MemoryDiagnoser]
public class ExcelLoadedValueReadBenchmarks {
    [GlobalSetup]
    public void Setup() {
        global::OfficeIMO.Benchmarks.BenchmarkProcessPriority.ApplyConfiguredPriority();
        XlBenchShapeScenarios.ValidateReadAll();
    }

    [Benchmark(Baseline = true)]
    public long OfficeIMO() => XlBenchShapeScenarios.OfficeImoReadAll();

    [Benchmark]
    public long OpenXml() => XlBenchShapeScenarios.OpenXmlReadAll();
}
