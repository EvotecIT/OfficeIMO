using System.Data;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Measures complete exports with the public defaults, including shared strings
/// and explicit cell references. The compact streaming profile has a separate lane.
/// </summary>
[MemoryDiagnoser]
public class ExcelDefaultTextWriteBenchmarks {
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
        ExcelTextWriteFixture.Validate(_table, nameof(OfficeIMO), WritePackage());
    }

    [Benchmark]
    public int OfficeIMO() => WritePackage().Length;

    private byte[] WritePackage() {
        using var stream = new MemoryStream();
        using var reader = _table.CreateDataReader();
        ExcelDocument.WriteDataReader(stream, reader);
        return stream.ToArray();
    }
}
