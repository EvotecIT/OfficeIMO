using System.Text;
using BenchmarkDotNet.Attributes;

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>
/// Compares the quote-free TSV profile with default CSV parsing on the same values.
/// Both inputs contain no quotes and differ only in the delimiter, so the gap is the
/// cost of quote handling. Setup checks that both dialects decode identical fields.
/// </summary>
[MemoryDiagnoser]
public class CsvTsvDialectBenchmarks
{
    private string _csvText = null!;
    private string _tsvText = null!;
    private string _csvPath = null!;
    private string _tsvPath = null!;
    private CsvReaderOptions _csvOptions = null!;
    private CsvReaderOptions _tsvOptions = null!;
    private long _expectedChecksum;

    [Params(100_000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        var csv = new StringBuilder();
        var tsv = new StringBuilder();
        string[] header = { "Id", "Name", "City", "Amount", "Date", "Code", "Flag", "Ratio", "Notes", "Tail" };
        csv.AppendLine(string.Join(",", header));
        tsv.AppendLine(string.Join("\t", header));
        for (int i = 0; i < RowCount; i++)
        {
            string[] row = {
                i.ToString(System.Globalization.CultureInfo.InvariantCulture), "Name " + (i % 997), "City" + (i % 31),
                (i * 1.25m).ToString(System.Globalization.CultureInfo.InvariantCulture), "2026-11-" + (1 + i % 28).ToString("00", System.Globalization.CultureInfo.InvariantCulture),
                "C" + (i % 4093), (i % 2 == 0) ? "true" : "false", (i / 7.0).ToString("0.000", System.Globalization.CultureInfo.InvariantCulture),
                "Plain notes without delimiters for row " + i, "end"
            };
            csv.AppendLine(string.Join(",", row));
            tsv.AppendLine(string.Join("\t", row));
        }

        _csvText = csv.ToString();
        _tsvText = tsv.ToString();
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-tsv-bench-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        _csvPath = Path.Combine(directory, "data.csv");
        _tsvPath = Path.Combine(directory, "data.tsv");
        File.WriteAllText(_csvPath, _csvText, new UTF8Encoding(false));
        File.WriteAllText(_tsvPath, _tsvText, new UTF8Encoding(false));
        _csvOptions = new CsvReaderOptions();
        _tsvOptions = CsvProfiles.CreateReaderOptions(CsvProfile.Tsv);

        _expectedChecksum = Csv_TextReader();
        Validate(Tsv_TextReader(), nameof(Tsv_TextReader));
        Validate(Csv_PathReader(), nameof(Csv_PathReader));
        Validate(Tsv_PathReader(), nameof(Tsv_PathReader));
        long csvSpans = Csv_FieldSpans();
        if (Tsv_FieldSpans() != csvSpans) throw new InvalidOperationException("TSV field spans differ from CSV field spans.");
    }

    [GlobalCleanup]
    public void Cleanup() => Directory.Delete(Path.GetDirectoryName(_csvPath)!, recursive: true);

    [Benchmark(Baseline = true)]
    public long Csv_TextReader() => ReadAll(CsvDocument.OpenReader(new StringReader(_csvText), _csvOptions));

    [Benchmark]
    public long Tsv_TextReader() => ReadAll(CsvDocument.OpenReader(new StringReader(_tsvText), _tsvOptions));

    [Benchmark]
    public long Csv_PathReader() => ReadAll(CsvDocument.OpenReader(_csvPath, _csvOptions));

    [Benchmark]
    public long Tsv_PathReader() => ReadAll(CsvDocument.OpenReader(_tsvPath, _tsvOptions));

    [Benchmark]
    public long Csv_FieldSpans() => SumSpans(_csvText, _csvOptions);

    [Benchmark]
    public long Tsv_FieldSpans() => SumSpans(_tsvText, _tsvOptions);

    private static long ReadAll(CsvDataReader reader)
    {
        using (reader)
        {
            long checksum = 0;
            while (reader.Read())
            {
                for (int i = 0; i < reader.FieldCount; i++) checksum += reader.GetString(i).Length;
                checksum++;
            }

            return checksum;
        }
    }

    private static long SumSpans(string text, CsvReaderOptions options)
    {
        long checksum = 0;
        CsvDocument.ReadFieldSpansFromText(text, (record, field, value) => checksum += value.Length + record + field, options);
        return checksum;
    }

    private void Validate(long actual, string name)
    {
        if (actual != _expectedChecksum)
            throw new InvalidOperationException($"{name} checksum {actual} differs from {_expectedChecksum}.");
    }
}
