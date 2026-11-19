using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>Materializes and formats the same 50,000-by-15 logical values through loaded workbook APIs.</summary>
[MemoryDiagnoser]
public class ExcelLoadedValueReadBenchmarks {
    [GlobalSetup]
    public void Setup() {
        global::OfficeIMO.Benchmarks.BenchmarkProcessPriority.ApplyConfiguredPriority();
        // ClosedXML reads the fixture as the oracle; both lanes must return every value in position.
        ExcelOutputValidation.Check(["OfficeIMO", "Open XML SDK"], "Default",
            "50,000 x 15 values compared one by one with a ClosedXML read of the same workbook",
            XlBenchShapeScenarios.ValidateReadAll);
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("case:Default")]
    public long OfficeIMO() => XlBenchShapeScenarios.OfficeImoReadAll();

    [Benchmark]
    [BenchmarkCategory("case:Default")]
    public long OpenXml() => XlBenchShapeScenarios.OpenXmlReadAll();
}
