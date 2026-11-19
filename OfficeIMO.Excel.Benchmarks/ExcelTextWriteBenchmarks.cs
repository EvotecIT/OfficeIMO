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
        string detail = $"TextLength={TextLength}; TextShape={TextShape}; ExcelDataReader reread the headers and every cell";
        ExcelOutputValidation.Check("OfficeIMO", "Default", detail,
            () => ExcelTextWriteFixture.Validate(_table, nameof(OfficeIMO), ExcelLibraryComparisonRunner.OfficeImoWriteDataReaderCompactPackageBytes(_table)));
        ExcelOutputValidation.Check("SpreadCheetah", "Default", detail,
            () => ExcelTextWriteFixture.Validate(_table, nameof(SpreadCheetah), ExcelLibraryComparisonRunner.SpreadCheetahWriteDataReaderPlainBytes(_table)));
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("case:Default")]
    public int OfficeIMO() => ExcelLibraryComparisonRunner.OfficeImoWriteDataReaderCompactPackageBytes(_table).Length;

    [Benchmark]
    [BenchmarkCategory("case:Default")]
    public int SpreadCheetah() => ExcelLibraryComparisonRunner.SpreadCheetahWriteDataReaderPlainBytes(_table).Length;

}
