using System.Data;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>Measures complete XLSX exports of short labels and long text, with independent output validation.</summary>
[MemoryDiagnoser]
public class ExcelTextWriteBenchmarks {
    private DataTable _table = null!;

    [Params(64, 4096)]
    public int TextLength { get; set; }

    [Params("Plain", "Escaped", "Markup")]
    public string TextShape { get; set; } = "Plain";

    [GlobalSetup]
    public void Setup() {
        string? priority = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_PROCESS_PRIORITY");
        if (!string.IsNullOrEmpty(priority)) BenchmarkProcessorAffinity.ApplyPriority(priority);
        _table = ExcelTextWriteFixture.Create(TextLength, TextShape);
        ExcelTextWriteFixture.Validate(_table, nameof(OfficeIMO), ExcelLibraryComparisonRunner.OfficeImoWriteDataReaderCompactPackageBytes(_table));
        ExcelTextWriteFixture.Validate(_table, nameof(SpreadCheetah), ExcelLibraryComparisonRunner.SpreadCheetahWriteDataReaderPlainBytes(_table));
    }

    [Benchmark(Baseline = true)]
    public int OfficeIMO() => ExcelLibraryComparisonRunner.OfficeImoWriteDataReaderCompactPackageBytes(_table).Length;

    [Benchmark]
    public int SpreadCheetah() => ExcelLibraryComparisonRunner.SpreadCheetahWriteDataReaderPlainBytes(_table).Length;

}
