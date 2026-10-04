using System.Globalization;
using System.IO.Compression;
using System.Text;
using BenchmarkDotNet.Attributes;
using CsvHelper.Configuration;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>
/// Measures complete document saves to fresh caller-owned memory streams.
/// This is an OfficeIMO API comparison, not a cross-library parity lane.
/// </summary>
[MemoryDiagnoser]
public class CsvDocumentSaveBenchmarks {
    private static readonly string[] Headers = ["Id", "Label", "Notes", "Enabled", "Score"];
    private static readonly Encoding Utf8 = new UTF8Encoding(false, true);
    private CsvDocument _document = null!;
    private CsvSaveOptions _options = null!;
    private object?[][] _rows = [];

    [Params(1000, 25000)]
    public int RowCount { get; set; }

    [Params("Plain", "Quoted", "MixedJson")]
    public string Shape { get; set; } = "Plain";

    [Params(CsvCompressionType.None, CsvCompressionType.GZip, CsvCompressionType.Deflate,
        CsvCompressionType.Brotli, CsvCompressionType.ZLib)]
    public CsvCompressionType Compression { get; set; }

    [GlobalSetup]
    public async Task Setup() {
        string? priority = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_PROCESS_PRIORITY");
        if (!string.IsNullOrEmpty(priority)) BenchmarkProcessorAffinity.ApplyPriority(priority);
        if (Shape is not ("Plain" or "Quoted" or "MixedJson"))
            throw new ArgumentOutOfRangeException(nameof(Shape));

        _options = new CsvSaveOptions {
            NewLine = "\n", Encoding = Utf8, CompressionType = Compression,
            CompressionLevel = CompressionLevel.Fastest
        };
        _document = new CsvDocument().WithHeader(Headers);
        _rows = new object?[RowCount][];
        for (int index = 0; index < RowCount; index++) {
            string text = Shape switch {
                "Plain" => "ordinary description " + index.ToString(CultureInfo.InvariantCulture),
                "Quoted" => "Łódź 🚀, \"row " + index.ToString(CultureInfo.InvariantCulture) + "\"\nnext line",
                _ => "{\"row\":" + index.ToString(CultureInfo.InvariantCulture)
                    + ",\"city\":\"Łódź 🚀 漢字\",\"note\":\"quoted value\"}"
            };
            object?[] row = [index, "row " + index.ToString(CultureInfo.InvariantCulture),
                index % 11 == 0 ? null : text, (index & 1) == 0, index * 1.25m];
            _rows[index] = row;
            _document.AddRow(row);
        }

        string expected = CreateReferenceText();
        using var sync = new MemoryStream();
        _document.Save(sync, _options);
        Validate(nameof(Save), sync, expected);
        using var asyncOutput = new MemoryStream();
        await _document.SaveAsync(asyncOutput, _options).ConfigureAwait(false);
        Validate(nameof(SaveAsync), asyncOutput, expected);
        Console.WriteLine($"Validated document save {Shape}/{Compression}: {RowCount} rows; sync={sync.Length}, async={asyncOutput.Length} bytes.");
    }

    [Benchmark(Baseline = true)]
    public long Save() {
        using var output = new MemoryStream();
        _document.Save(output, _options);
        return output.Length;
    }

    [Benchmark]
    public async Task<long> SaveAsync() {
        using var output = new MemoryStream();
        await _document.SaveAsync(output, _options).ConfigureAwait(false);
        return output.Length;
    }

    private string CreateReferenceText() {
        using var output = new StringWriter(CultureInfo.InvariantCulture);
        using (var csv = new global::CsvHelper.CsvWriter(output,
                   new CsvConfiguration(CultureInfo.InvariantCulture) { NewLine = "\n" }, leaveOpen: true)) {
            foreach (string header in Headers) csv.WriteField(header);
            csv.NextRecord();
            foreach (object?[] row in _rows) {
                foreach (object? value in row) csv.WriteField(value);
                csv.NextRecord();
            }
        }
        return output.ToString();
    }

    private void Validate(string method, MemoryStream output, string expected) {
        if (!output.CanWrite || output.Length == 0)
            throw new InvalidDataException($"{method} must leave a nonempty caller-owned stream open.");
        output.Position = 0;
        using Stream decoded = Compression switch {
            CsvCompressionType.None => new MemoryStream(output.ToArray(), writable: false),
            CsvCompressionType.GZip => new GZipStream(output, CompressionMode.Decompress, leaveOpen: true),
            CsvCompressionType.Deflate => new DeflateStream(output, CompressionMode.Decompress, leaveOpen: true),
            CsvCompressionType.Brotli => new BrotliStream(output, CompressionMode.Decompress, leaveOpen: true),
            CsvCompressionType.ZLib => new ZLibStream(output, CompressionMode.Decompress, leaveOpen: true),
            _ => throw new ArgumentOutOfRangeException(nameof(Compression))
        };
        using var reader = new StreamReader(decoded, Utf8, detectEncodingFromByteOrderMarks: false);
        string text = reader.ReadToEnd();
        if (!string.Equals(text, expected, StringComparison.Ordinal))
            throw new InvalidDataException($"{method} differs from the independently formatted reference.");
        CsvBenchmarkOutputValidator.Validate(method, text, Headers, RowCount,
            expectedTextRows: null, expectedObjectRows: _rows);
    }
}
