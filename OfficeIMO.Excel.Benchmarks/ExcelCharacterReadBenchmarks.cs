using System.Data.Common;
using System.Globalization;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Consumes every UTF-16 character through GetString or sixteen-character GetChars calls.
/// Short values use 5,000 rows of 64 characters; long values use 128 rows of 4,096 characters.
/// Every variant reads the same cached workbook. Setup checks all values and both access
/// modes against the generated source rows; each measured operation checks the ordered hash.
/// </summary>
[MemoryDiagnoser]
public class ExcelCharacterReadBenchmarks {
    private readonly char[] _chunk = new char[16];
    private byte[] _input = [];
    private string[] _values = [];
    private ulong _expected;

    [Params(ExcelFileFormat.Xlsx, ExcelFileFormat.Xlsb, ExcelFileFormat.Xls)]
    public ExcelFileFormat Format { get; set; }

    [Params("Short", "Long")]
    public string Shape { get; set; } = "Short";

    [GlobalSetup]
    public void Setup() {
        BenchmarkProcessorAffinity.ApplyPriority(
            Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_PROCESS_PRIORITY") ?? "Normal");
        if (Shape is not ("Short" or "Long")) throw new ArgumentOutOfRangeException(nameof(Shape));
        int rows = Shape == "Short" ? 5_000 : 128;
        int length = Shape == "Short" ? 64 : 4_096;
        _values = Enumerable.Range(0, rows).Select(index => CreateText(index, length)).ToArray();
        ulong expected = 14695981039346656037UL;
        for (int row = 0; row < _values.Length; row++) {
            foreach (char value in _values[row]) expected = Append(expected, value);
            expected = Append(expected, (uint)row);
        }
        _expected = expected;

        string root = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_DATA")
            ?? Path.Combine(Path.GetTempPath(), "OfficeIMO", "Benchmarks", "Characters");
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, $"characters-{rows}-{length}.{Format.ToString().ToLowerInvariant()}");
        if (!File.Exists(path)) CreateFixture(path);
        _input = File.ReadAllBytes(path);
        Validate(ReadCore(chunks: false, validate: true));
        Validate(ReadCore(chunks: true, validate: true));
    }

    [Benchmark]
    public ulong StringValues() => Validate(ReadCore(chunks: false, validate: false));

    [Benchmark]
    public ulong CharacterChunks() => Validate(ReadCore(chunks: true, validate: false));

    private ulong Validate(ulong actual) {
        if (actual != _expected) throw new InvalidDataException("The character-read checksum differs.");
        return actual;
    }

    private ulong ReadCore(bool chunks, bool validate) {
        using DbDataReader reader = ExcelDocument.OpenReader(_input);
        if (reader.FieldCount != 2 || reader.GetName(0) != "Text" || reader.GetName(1) != "Id")
            throw new InvalidDataException("The character-read schema differs.");
        ulong hash = 14695981039346656037UL;
        int row = 0;
        while (reader.Read()) {
            if (row >= _values.Length) throw new InvalidDataException("The character-read row count differs.");
            string expected = _values[row];
            if (chunks) {
                long total = reader.GetChars(0, 0, null, 0, 0);
                if (total != expected.Length) throw new InvalidDataException("The character-read length differs.");
                long offset = 0;
                long copied;
                while ((copied = reader.GetChars(0, offset, _chunk, 0, _chunk.Length)) != 0) {
                    int count = checked((int)copied);
                    if (validate && !expected.AsSpan(checked((int)offset), count).SequenceEqual(_chunk.AsSpan(0, count)))
                        throw new InvalidDataException($"Character range at row {row} differs.");
                    for (int index = 0; index < count; index++) hash = Append(hash, _chunk[index]);
                    offset += count;
                }
                if (offset != total) throw new InvalidDataException("The copied character count differs.");
            } else {
                string value = reader.GetString(0);
                if (validate && !string.Equals(value, expected, StringComparison.Ordinal))
                    throw new InvalidDataException($"String at row {row} differs.");
                foreach (char character in value) hash = Append(hash, character);
            }
            int id = reader.GetInt32(1);
            if (id != row) throw new InvalidDataException("The character-read row order differs.");
            hash = Append(hash, (uint)id);
            row++;
        }
        if (row != _values.Length) throw new InvalidDataException("The character-read row count differs.");
        return hash;
    }

    private void CreateFixture(string path) {
        string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try {
            using (var writer = ExcelDocument.CreateWriter(Format, temporary,
                new ExcelWriterOptions { CreateTable = false, UseSharedStrings = false })) {
                writer.WriteTable("Data", Enumerable.Range(0, _values.Length), new[] { "Text", "Id" },
                    (row, index) => { row.Write(_values[index]); row.Write(index); });
            }
            File.Move(temporary, path);
        } finally {
            if (File.Exists(temporary)) File.Delete(temporary);
        }
    }

    private static string CreateText(int row, int length) {
        string prefix = row.ToString("D5", CultureInfo.InvariantCulture) + "|Żółw🐢|";
        return prefix + new string('x', length - prefix.Length);
    }

    private static ulong Append(ulong hash, uint value) => unchecked((hash ^ value) * 1099511628211UL);
}
