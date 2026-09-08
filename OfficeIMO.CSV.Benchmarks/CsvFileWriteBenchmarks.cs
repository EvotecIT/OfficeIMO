using System.Globalization;
using System.Text;
using BenchmarkDotNet.Attributes;
using CsvHelper.Configuration;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>
/// Writes and closes UTF-8 files with identical stream buffers and quoting.
/// This includes file creation and flush to the operating system, without a
/// durable-storage flush. Setup validates bytes and every decoded field.
/// </summary>
[MemoryDiagnoser]
public class CsvFileWriteBenchmarks {
    private static readonly string[] Headers = ["Id", "Label", "Notes", "Tail"];
    private static readonly Type[] Types = [typeof(string), typeof(string), typeof(string), typeof(string)];
    private static readonly Encoding Utf8 = new UTF8Encoding(false, true);
    private object?[][] _rows = [];
    private string[] _headers = Headers;
    private Type[] _types = Types;
    private bool _typedValues;
    private string _delimiter = ",";
    private string? _directory;
    private string _officePath = "";
    private string _peerPath = "";

    [Params("ShortAscii", "ShortUnicode", "DenseJson", "QuoteRuns", "LongNotes", "TypedValues")]
    public string Shape { get; set; } = "ShortAscii";

    [Params(CsvQuoteMode.AsNeeded, CsvQuoteMode.Always)]
    public CsvQuoteMode QuoteMode { get; set; }

    [GlobalSetup]
    public void Setup() {
        string? priority = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_PROCESS_PRIORITY");
        if (!string.IsNullOrEmpty(priority)) BenchmarkProcessorAffinity.ApplyPriority(priority);
        _typedValues = Shape == "TypedValues";
        _headers = Headers;
        _types = Types;
        (string delimiter, string payload) = Shape switch {
            "ShortAscii" => (",", "short ordinary description"),
            "ShortUnicode" => ("||", "Łódź 🚀 漢字 || \"notes\"\nnext line"),
            "DenseJson" => (",", "{" + string.Concat(Enumerable.Repeat("\"key\":\"value\",", 292)) + "\"last\":null}"),
            "QuoteRuns" => ("※", new string('"', 4096)),
            "LongNotes" => (";", new string('n', 2048) + ";\"Łódź 🚀\"\n" + new string('x', 2048)),
            "TypedValues" => (";", ""),
            _ => throw new InvalidOperationException($"Unknown file-write shape: {Shape}.")
        };
        _delimiter = delimiter;
        if (_typedValues) {
            _headers = ["Id", "Label", "Score", "Created", "Enabled"];
            _types = [typeof(int), typeof(string), typeof(decimal), typeof(DateTime), typeof(bool)];
            var start = new DateTime(2026, 1, 1, 12, 30, 0, DateTimeKind.Utc);
            _rows = Enumerable.Range(0, 1000).Select(index => new object?[] {
                index, index % 11 == 0 ? null : "Łódź; \"row " + index.ToString(CultureInfo.InvariantCulture) + "\"",
                index * 1.25m, start.AddMinutes(index), (index & 1) == 0
            }).ToArray();
        } else {
            _rows = Enumerable.Range(0, 1000).Select(index => new object?[] {
                index.ToString(CultureInfo.InvariantCulture), "row " + index.ToString(CultureInfo.InvariantCulture),
                index.ToString(CultureInfo.InvariantCulture) + payload, "end-" + index.ToString(CultureInfo.InvariantCulture)
            }).ToArray();
        }

        string root = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_OUTPUT")
            ?? Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.Benchmarks");
        _directory = Path.Combine(root, "file-write-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_directory);
        _officePath = Path.Combine(_directory, "officeimo.csv");
        _peerPath = Path.Combine(_directory, "csvhelper.csv");
        try {
            OfficeIMO();
            CsvHelper();
            byte[] office = File.ReadAllBytes(_officePath);
            byte[] peer = File.ReadAllBytes(_peerPath);
            if (!office.AsSpan().SequenceEqual(peer)) throw new InvalidOperationException("File writers produced different UTF-8 bytes.");
            CsvBenchmarkOutputValidator.Validate(nameof(OfficeIMO), Utf8.GetString(office), _headers, _rows.Length,
                expectedTextRows: null, expectedObjectRows: _rows, delimiter: _delimiter);
            CsvBenchmarkOutputValidator.Validate(nameof(CsvHelper), Utf8.GetString(peer), _headers, _rows.Length,
                expectedTextRows: null, expectedObjectRows: _rows, delimiter: _delimiter);
            Console.WriteLine($"Validated CSV file {Shape}/{QuoteMode}: {_rows.Length} rows, {office.Length} identical UTF-8 bytes.");
        } catch {
            Cleanup();
            throw;
        }
    }

    [Benchmark(Baseline = true)]
    public int OfficeIMO() {
        using var stream = new FileStream(_officePath, FileMode.Create, FileAccess.Write, FileShare.None, 65536);
        using (var writer = new StreamWriter(stream, Utf8, 65536, leaveOpen: true)) {
            using var reader = new BenchmarkArrayDataReader(_headers, _rows, _types);
            CsvDocument.WriteDataReader(writer, reader,
                new CsvSaveOptions { NewLine = "\n", DelimiterText = _delimiter, QuoteMode = QuoteMode,
                    DateTimeFormat = _typedValues ? "O" : null });
        }
        return checked((int)stream.Position);
    }

    [Benchmark]
    public int CsvHelper() {
        using var stream = new FileStream(_peerPath, FileMode.Create, FileAccess.Write, FileShare.None, 65536);
        using (var writer = new StreamWriter(stream, Utf8, 65536, leaveOpen: true)) {
            var configuration = new CsvConfiguration(CultureInfo.InvariantCulture) { NewLine = "\n", Delimiter = _delimiter };
            if (QuoteMode == CsvQuoteMode.Always) configuration.ShouldQuote = _ => true;
            using var csv = new global::CsvHelper.CsvWriter(writer, configuration, leaveOpen: true);
            if (_typedValues) csv.Context.TypeConverterOptionsCache.GetOptions<DateTime>().Formats = ["O"];
            using var reader = new BenchmarkArrayDataReader(_headers, _rows, _types);
            for (int column = 0; column < reader.FieldCount; column++) csv.WriteField(reader.GetName(column));
            csv.NextRecord();
            while (reader.Read()) {
                if (_typedValues) {
                    csv.WriteField(reader.GetInt32(0));
                    csv.WriteField(reader.GetString(1));
                    csv.WriteField(reader.GetDecimal(2));
                    csv.WriteField(reader.GetDateTime(3));
                    csv.WriteField(reader.GetBoolean(4));
                } else {
                    for (int column = 0; column < reader.FieldCount; column++) csv.WriteField(reader.GetString(column));
                }
                csv.NextRecord();
            }
        }
        return checked((int)stream.Position);
    }

    [GlobalCleanup]
    public void Cleanup() {
        if (_directory == null) return;
        File.Delete(_officePath);
        File.Delete(_peerPath);
        Directory.Delete(_directory);
        _directory = null;
    }
}
