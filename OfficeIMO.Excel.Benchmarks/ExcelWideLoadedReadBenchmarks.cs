using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Supplemental 10,000-by-150 loaded-sheet presence scan. This is a separate wide
/// workload, not one of XLBench's published scenarios. Both APIs include headers.
/// </summary>
[MemoryDiagnoser]
public class ExcelWideLoadedReadBenchmarks {
    [GlobalSetup]
    public void Setup() {
        // Each lane throws unless it counts exactly 10,000 x 150 non-empty cells.
        ExcelOutputValidation.Check("OfficeIMO", "Default", "non-empty cell count equals 10,000 x 150", () => OfficeIMO());
        ExcelOutputValidation.Check("Open XML SDK", "Default", "non-empty cell count equals 10,000 x 150", () => OpenXml());
    }

    [Benchmark]
    [BenchmarkCategory("case:Default")]
    public long OfficeIMO() => XlBenchShapeScenarios.OfficeImoWideRead();

    [Benchmark]
    [BenchmarkCategory("case:Default")]
    public long OpenXml() => XlBenchShapeScenarios.OpenXmlWideRead();
}
