using System.Globalization;
using BenchmarkDotNet.Attributes;
using CsvHelper.Configuration;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>Exports short labels and long notes with embedded quotes through public CSV writers.</summary>
[MemoryDiagnoser]
public class CsvTextWriteBenchmarks {
    private static readonly string[] Headers = ["Id", "Notes"];
    private object?[][] _rows = [];

    [Params(64, 4096)]
    public int TextLength { get; set; }

    [Params(CsvQuoteMode.AsNeeded, CsvQuoteMode.Always)]
    public CsvQuoteMode QuoteMode { get; set; }

    [Params("Notes", "Json", "Quotes")]
    public string TextShape { get; set; } = "Notes";

    [GlobalSetup]
    public void Setup() {
        string? priority = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_PROCESS_PRIORITY");
        if (!string.IsNullOrEmpty(priority)) BenchmarkProcessorAffinity.ApplyPriority(priority);
        string payload = TextShape switch {
            "Notes" => ",\"" + new string('n', TextLength / 2) + "\" Łódź\n" + new string('x', TextLength / 2),
            "Json" => "{" + string.Concat(Enumerable.Repeat("\"key\":\"value\",", Math.Max(1, TextLength / 14))) + "\"last\":null}",
            "Quotes" => new string('"', TextLength),
            _ => throw new InvalidOperationException($"Unknown text shape: {TextShape}.")
        };
        _rows = Enumerable.Range(0, 1000).Select(index => new object?[] {
            index.ToString(CultureInfo.InvariantCulture),
            index.ToString(CultureInfo.InvariantCulture) + payload
        }).ToArray();
        using var officeWriter = new StringWriter(CultureInfo.InvariantCulture);
        using var peerWriter = new StringWriter(CultureInfo.InvariantCulture);
        WriteOfficeIMO(officeWriter);
        WriteCsvHelper(peerWriter);
        CsvBenchmarkOutputValidator.Validate(nameof(OfficeIMO), officeWriter.ToString(), Headers, _rows.Length, expectedTextRows: null, expectedObjectRows: _rows);
        CsvBenchmarkOutputValidator.Validate(nameof(CsvHelper), peerWriter.ToString(), Headers, _rows.Length, expectedTextRows: null, expectedObjectRows: _rows);
        if (officeWriter.ToString() != peerWriter.ToString()) throw new InvalidOperationException("CSV writers produced different text or quoting.");
    }

    [Benchmark(Baseline = true)]
    public int OfficeIMO() {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        WriteOfficeIMO(writer);
        return writer.GetStringBuilder().Length;
    }

    [Benchmark]
    public int CsvHelper() {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        WriteCsvHelper(writer);
        return writer.GetStringBuilder().Length;
    }

    private void WriteOfficeIMO(TextWriter writer) {
        using var reader = new BenchmarkArrayDataReader(Headers, _rows, [typeof(string), typeof(string)]);
        CsvDocument.WriteDataReader(writer, reader, new CsvSaveOptions { NewLine = "\n", QuoteMode = QuoteMode });
    }

    private void WriteCsvHelper(TextWriter writer) {
        var configuration = new CsvConfiguration(CultureInfo.InvariantCulture) { NewLine = "\n" };
        if (QuoteMode == CsvQuoteMode.Always) configuration.ShouldQuote = _ => true;
        using var csv = new global::CsvHelper.CsvWriter(writer, configuration, leaveOpen: true);
        using var reader = new BenchmarkArrayDataReader(Headers, _rows, [typeof(string), typeof(string)]);
        for (int column = 0; column < reader.FieldCount; column++) csv.WriteField(reader.GetName(column));
        csv.NextRecord();
        while (reader.Read()) {
            for (int column = 0; column < reader.FieldCount; column++) csv.WriteField(reader.GetString(column));
            csv.NextRecord();
        }
    }
}
