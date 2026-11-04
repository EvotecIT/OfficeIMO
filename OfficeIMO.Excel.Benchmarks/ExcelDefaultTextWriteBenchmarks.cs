using System.Data;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Measures complete bounded exports with inline text and explicit cell references.
/// The compact profile omits data cell references in a separate lane.
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
        var destination1 = stream;

        var writeOptions1 = new ExcelWriterOptions {
            UseSharedStrings = false,
            CreateTable = false
        };

        using (var workbookWriter1 = ExcelDocument.CreateWriter(destination1, writeOptions1, leaveOpen: true)) {
            workbookWriter1.WriteTable(writeOptions1.SheetName, reader);
        }

        if (destination1.CanSeek)
            destination1.Position = 0;
        return stream.ToArray();
    }
}
