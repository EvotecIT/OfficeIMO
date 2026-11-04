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
        OfficeIMO();
        OpenXml();
    }

    [Benchmark]
    public long OfficeIMO() => XlBenchShapeScenarios.OfficeImoWideRead();

    [Benchmark]
    public long OpenXml() => XlBenchShapeScenarios.OpenXmlWideRead();
}
