using System.Data;
using BenchmarkDotNet.Attributes;
using ExcelDataReader;
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
        _table = new DataTable();
        _table.Columns.Add("Id", typeof(string));
        _table.Columns.Add("Notes", typeof(string));
        for (int index = 0; index < 1000; index++) {
            string marker = TextShape == "Escaped" ? " & <tag> Łódź\n" : " plain Łódź ";
            string payload = TextShape == "Markup"
                ? string.Concat(Enumerable.Repeat("<tag>&data</tag>", Math.Max(1, TextLength / 16)))
                : new string('n', TextLength / 2) + marker + new string('x', TextLength / 2);
            _table.Rows.Add(index.ToString(System.Globalization.CultureInfo.InvariantCulture),
                index + payload);
        }
        Validate(nameof(OfficeIMO), ExcelLibraryComparisonRunner.OfficeImoWriteDataReaderCompactPackageBytes(_table));
        Validate(nameof(SpreadCheetah), ExcelLibraryComparisonRunner.SpreadCheetahWriteDataReaderPlainBytes(_table));
    }

    [Benchmark(Baseline = true)]
    public int OfficeIMO() => ExcelLibraryComparisonRunner.OfficeImoWriteDataReaderCompactPackageBytes(_table).Length;

    [Benchmark]
    public int SpreadCheetah() => ExcelLibraryComparisonRunner.SpreadCheetahWriteDataReaderPlainBytes(_table).Length;

    private void Validate(string method, byte[] bytes) {
        using var stream = new MemoryStream(bytes, writable: false);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        if (!reader.Read() || reader.FieldCount != 2 || reader.GetString(0) != "Id" || reader.GetString(1) != "Notes")
            throw new InvalidOperationException($"{method} header mismatch.");
        foreach (DataRow row in _table.Rows) {
            if (!reader.Read() || reader.GetString(0) != (string)row[0] || reader.GetString(1) != (string)row[1])
                throw new InvalidOperationException($"{method} row mismatch.");
        }
        if (reader.Read()) throw new InvalidOperationException($"{method} wrote extra rows.");
        Console.WriteLine($"Validated {method}: {_table.Rows.Count} rows, {bytes.Length} XLSX bytes.");
    }
}
