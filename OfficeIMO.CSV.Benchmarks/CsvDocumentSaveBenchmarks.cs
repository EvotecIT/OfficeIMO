using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using BenchmarkDotNet.Attributes;
using CsvHelper.Configuration;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>
/// Measures complete document saves to memory, bytes, and files.
/// This is an OfficeIMO API comparison, not a cross-library parity lane.
/// </summary>
[MemoryDiagnoser]
public class CsvDocumentSaveBenchmarks {
    private static readonly string[] Headers = ["Id", "Label", "Notes", "Enabled", "Score"];
    private static readonly Type[] FieldTypes = [typeof(int), typeof(string), typeof(string), typeof(bool), typeof(decimal)];
    private static readonly Encoding Utf8 = new UTF8Encoding(false, true);
    private CsvDocument _document = null!;
    private CsvSaveOptions _options = null!;
    private object?[][] _rows = [];
    private string? _directory;
    private string _file = null!, _asyncFile = null!, _rowFile = null!;
    private readonly Dictionary<string, long> _expectedOutputLengths = [];
    private int _rowWriterBufferSize;

    /// <summary>Identifies the independently formatted, uncompressed output validated by setup.</summary>
    public string ExpectedCsvSha256 { get; private set; } = "";

    /// <summary>Gets each operation's validated output length, including its compression framing.</summary>
    public IReadOnlyDictionary<string, long> ExpectedOutputLengths => _expectedOutputLengths;

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
        if (Shape is not ("Plain" or "Quoted" or "MixedJson" or "LongUnicode"))
            throw new ArgumentOutOfRangeException(nameof(Shape));
        // Resolve the loaded build's default during setup so snapshot comparisons
        // use identical benchmark IL while measuring newly compiled callers.
        _rowWriterBufferSize = (int)typeof(CsvRowWriter).GetMethod(nameof(CsvRowWriter.CreateFile))!
            .GetParameters()[3].DefaultValue!;

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
                "LongUnicode" => new string('x', 32767 + index % 3) + "🚀 漢字, \"end\"\r\nnext",
                _ => "{\"row\":" + index.ToString(CultureInfo.InvariantCulture)
                    + ",\"city\":\"Łódź 🚀 漢字\",\"note\":\"quoted value\"}"
            };
            object?[] row = [index, "row " + index.ToString(CultureInfo.InvariantCulture),
                index % 11 == 0 ? null : text, (index & 1) == 0, index * 1.25m];
            _rows[index] = row;
            _document.AddRow(row);
        }

        string expected = CreateReferenceText();
        ExpectedCsvSha256 = Convert.ToHexString(SHA256.HashData(Utf8.GetBytes(expected)));
        string textOutput = _document.ToString(_options);
        if (!string.Equals(expected, textOutput, StringComparison.Ordinal))
            throw new InvalidDataException("Document text differs from the independently formatted CSV.");
        _expectedOutputLengths[nameof(ToText)] = textOutput.Length;
        using var sync = new MemoryStream();
        _document.Save(sync, _options);
        Validate(nameof(Save), sync, expected);
        using var asyncOutput = new MemoryStream();
        await _document.SaveAsync(asyncOutput, _options).ConfigureAwait(false);
        Validate(nameof(SaveAsync), asyncOutput, expected);
        using var bytes = new MemoryStream(_document.ToBytes(_options), writable: false);
        Validate(nameof(ToBytes), bytes, expected, callerOwned: false);
        using var sequential = new MemoryStream();
        WriteDataReaderCore(sequential, parallel: false);
        Validate(nameof(WriteDataReader), sequential, expected);
        using var parallel = new MemoryStream();
        WriteDataReaderCore(parallel, parallel: true);
        Validate(nameof(WriteDataReaderParallel), parallel, expected);
        string root = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_OUTPUT") ?? Path.GetTempPath();
        _directory = Path.Combine(Path.GetFullPath(root), "OfficeIMO.CsvDocumentSave-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_directory);
        _file = Path.Combine(_directory, "sync.csv");
        _asyncFile = Path.Combine(_directory, "async.csv");
        _rowFile = Path.Combine(_directory, "rows.csv");
        try {
            _document.Save(_file, _options);
            ValidateFile(nameof(SaveFile), _file, expected);
            await _document.SaveAsync(_asyncFile, _options).ConfigureAwait(false);
            ValidateFile(nameof(SaveFileAsync), _asyncFile, expected);
            RowWriterFile();
            ValidateFile(nameof(RowWriterFile), _rowFile, expected);
        } catch {
            Cleanup();
            throw;
        }
        Console.WriteLine($"Validated document save {Shape}/{Compression}: {RowCount} rows; CSV SHA256={ExpectedCsvSha256}; lengths={string.Join(",", _expectedOutputLengths.Select(pair => pair.Key + "=" + pair.Value))}.");
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

    [Benchmark]
    public long ToBytes() => _document.ToBytes(_options).LongLength;

    [Benchmark]
    public long ToText() => _document.ToString(_options).Length;

    [Benchmark]
    public long SaveFile() {
        _document.Save(_file, _options);
        return new FileInfo(_file).Length;
    }

    [Benchmark]
    public async Task<long> SaveFileAsync() {
        await _document.SaveAsync(_asyncFile, _options).ConfigureAwait(false);
        return new FileInfo(_asyncFile).Length;
    }

    [Benchmark]
    public long RowWriterFile() {
        using (var writer = CsvRowWriter.CreateFile(_rowFile, _options, bufferSize: _rowWriterBufferSize)) {
            using var reader = new BenchmarkArrayDataReader(Headers, _rows, FieldTypes);
            writer.WriteDataReader(reader);
        }
        return new FileInfo(_rowFile).Length;
    }

    [Benchmark]
    public long WriteDataReader() {
        using var output = new MemoryStream();
        WriteDataReaderCore(output, parallel: false);
        return output.Length;
    }

    [Benchmark]
    public long WriteDataReaderParallel() {
        using var output = new MemoryStream();
        WriteDataReaderCore(output, parallel: true);
        return output.Length;
    }

    private void WriteDataReaderCore(Stream output, bool parallel) {
        using var reader = new BenchmarkArrayDataReader(Headers, _rows, FieldTypes);
        if (parallel) {
            CsvDocument.WriteDataReaderParallel(output, reader, _options,
                new CsvWriteParallelOptions { MaxDegreeOfParallelism = 4, BatchSize = 512 });
        } else {
            CsvDocument.WriteDataReader(output, reader, _options);
        }
    }

    [GlobalCleanup]
    public void Cleanup() {
        if (_directory == null) return;
        File.Delete(_file);
        File.Delete(_asyncFile);
        File.Delete(_rowFile);
        Directory.Delete(_directory);
        _directory = null;
    }

    private void ValidateFile(string method, string path, string expected) {
        using var output = new MemoryStream(File.ReadAllBytes(path), writable: false);
        Validate(method, output, expected, callerOwned: false);
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

    private void Validate(string method, MemoryStream output, string expected, bool callerOwned = true) {
        if ((callerOwned && !output.CanWrite) || output.Length == 0)
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
        _expectedOutputLengths[method] = output.Length;
    }
}
