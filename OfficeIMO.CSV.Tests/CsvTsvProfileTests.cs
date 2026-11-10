using System.Text;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvTsvProfileTests {
    public enum ReadRoute {
        Parse, LoadStream, LoadPath, LoadAsyncPath, ReaderText, ReaderPath, ReaderStream, ReaderAsyncPath,
        DocumentReader, RawRecords, FieldSpansText, FieldSpansReader, ParallelTextRows, ParallelAggregate
    }

    public static IEnumerable<object[]> Routes() =>
        from ReadRoute route in Enum.GetValues(typeof(ReadRoute))
        from int rows in new[] { 6, 4000 } // 4000 rows exceed the 64 Ki quote-probe and SIMD thresholds.
        from bool trim in new[] { false, true } // Trimming has its own SIMD record path on x64.
        select new object[] { route, rows, trim };

    [Theory]
    [MemberData(nameof(Routes))]
    public async Task TsvReadsQuotesAsDataOnEveryReadRoute(ReadRoute route, int rowCount, bool trim) {
        string[][] source = CreateRowsWithStrayQuotes(rowCount);
        string text = string.Join("\n", source.Select(row => string.Join("\t", row))) + "\n";
        string[][] expected = trim ? source.Select(row => row.Select(value => value.Trim(' ')).ToArray()).ToArray() : source;
        string[][] actual = await ReadAsync(route, text, trim);
        Assert.Equal(expected.Length, actual.Length);
        for (int i = 0; i < expected.Length; i++) Assert.Equal(expected[i], actual[i]);
    }

    [Theory]
    [InlineData("a\tb")]
    [InlineData("a\nb")]
    [InlineData("a\rb")]
    public void TsvSaveRejectsFieldsItCannotRepresent(string value) {
        var document = CsvDocument.Create().WithHeader("Id", "Value").AddRow(1, value);
        CsvSaveOptions options = CsvProfiles.CreateSaveOptions(CsvProfile.Tsv);
        Assert.Throws<CsvException>(() => document.ToString(options));
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".tsv");
        try {
            Assert.Throws<CsvException>(() => document.Save(path, options));
            Assert.Throws<CsvException>(() => document.Save(path));
        } finally { File.Delete(path); }
        string quotes = CsvDocument.Create().WithHeader("Id", "Value").AddRow(1, "say \"hi\"").ToString(options);
        Assert.Equal("Id\tValue\n1\tsay \"hi\"\n", quotes);
    }

    [Fact]
    public void BackslashEscapesRoundTripPostgresCopyText() {
        var document = CsvDocument.Create().WithHeader("Id", "Value")
            .AddRow(1, "tab\there")
            .AddRow(2, "line\nbreak\r\nend")
            .AddRow(3, "C:\\path \\N \"quoted\"");
        CsvSaveOptions save = CsvProfiles.CreateSaveOptions(CsvProfile.Tsv);
        save.EscapeMode = CsvEscapeMode.Backslash;
        string text = document.ToString(save);
        Assert.Equal("Id\tValue\n1\ttab\\there\n2\tline\\nbreak\\r\\nend\n3\tC:\\\\path \\\\N \"quoted\"\n", text);

        CsvLoadOptions load = CsvProfiles.CreateLoadOptions(CsvProfile.Tsv);
        load.EscapeMode = CsvEscapeMode.Backslash;
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".tsv");
        try {
            File.WriteAllText(path, text + "4\t\\N unknown \\x stays\n");
            var reader = CsvProfiles.CreateReaderOptions(CsvProfile.Tsv);
            reader.EscapeMode = CsvEscapeMode.Backslash;
            using CsvDataReader rows = CsvDocument.OpenReader(path, reader);
            var values = new List<string>();
            while (rows.Read()) values.Add(rows.GetString(1));
            Assert.Equal(new[] { "tab\there", "line\nbreak\r\nend", "C:\\path \\N \"quoted\"", "\\N unknown \\x stays" }, values);
            CsvDocument loaded = CsvDocument.Parse(text, load);
            Assert.Equal(text, loaded.ToString(save));
        } finally { File.Delete(path); }

        load.QuoteParsingMode = CsvQuoteParsingMode.Strict;
        Assert.Throws<ArgumentException>(() => CsvDocument.Parse(text, load));
        save.QuoteMode = CsvQuoteMode.AsNeeded;
        Assert.Throws<ArgumentException>(() => document.ToString(save));
    }

    [Theory]
    [InlineData(".tsv")]
    [InlineData(".tab")]
    [InlineData(".tsv.gz")]
    public void TabSeparatedPathsDefaultToTsvWhenNoOptionsAreGiven(string extension) {
        string directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string source = Path.Combine(directory, "source" + extension);
            CsvDocument.Create().WithHeader("Id", "Note").AddRow(1, "\"open").AddRow(2, "a,b").Save(source);
            CsvDocument loaded = CsvDocument.Load(source);
            Assert.Equal(new[] { "\"open", "a,b" }, loaded.AsEnumerable().Select(row => row.AsString("Note")).ToArray());
            using (CsvDataReader reader = CsvDocument.OpenReader(source)) {
                Assert.True(reader.Read());
                Assert.Equal("\"open", reader.GetString(1));
            }
            if (extension == ".tsv") {
                Assert.Equal("Id\tNote\n1\t\"open\n2\ta,b\n", File.ReadAllText(source));
                string utf16 = Path.Combine(directory, "utf16.tsv");
                loaded.WithEncoding(Encoding.Unicode).Save(utf16);
                Assert.Equal(Encoding.Unicode.GetPreamble(), File.ReadAllBytes(utf16).Take(2).ToArray());
                Assert.Equal("\"open", CsvDocument.Load(utf16).AsEnumerable().First().AsString("Note"));
            }
        } finally { Directory.Delete(directory, recursive: true); }
    }

    private static string[][] CreateRowsWithStrayQuotes(int rowCount) {
        string[] notes = { "\"quoted\"", "he said \"hi", "\"", "\"\"", "mid\"dle with a long tail past thirty two chars", "plain" };
        var rows = new List<string[]> { new[] { "Id", "Note", "Tail" } };
        for (int i = 1; i < rowCount; i++)
            rows.Add(new[] {
                i.ToString(System.Globalization.CultureInfo.InvariantCulture),
                i % 3 == 0 ? "  " + notes[i % notes.Length] + " " : notes[i % notes.Length],
                "\"opens" + i });
        return rows.ToArray();
    }

    private static async Task<string[][]> ReadAsync(ReadRoute route, string text, bool trim) {
        CsvLoadOptions load = CsvProfiles.CreateLoadOptions(CsvProfile.Tsv);
        load.HasHeaderRow = false;
        load.TrimWhitespace = trim;
        CsvReaderOptions reader = CsvProfiles.CreateReaderOptions(CsvProfile.Tsv);
        reader.HasHeaderRow = false;
        reader.TrimWhitespace = trim;
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".tsv");
        File.WriteAllText(path, text);
        try {
            switch (route) {
                case ReadRoute.Parse: return Rows(CsvDocument.Parse(text, load));
                case ReadRoute.LoadStream: using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(text))) return Rows(CsvDocument.Load(stream, load));
                case ReadRoute.LoadPath: return Rows(CsvDocument.Load(path, load));
                case ReadRoute.LoadAsyncPath: return Rows(await CsvDocument.LoadAsync(path, load));
                case ReadRoute.ReaderText: using (var r = CsvDocument.OpenReader(new StringReader(text), reader)) return Rows(r);
                case ReadRoute.ReaderPath: using (var r = CsvDocument.OpenReader(path, reader)) return Rows(r);
                case ReadRoute.ReaderStream:
                    using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(text)))
                    using (var r = CsvDocument.OpenReader(stream, reader)) return Rows(r);
                case ReadRoute.DocumentReader: using (var r = CsvDocument.Parse(text, load).CreateDataReader(reader)) return Rows(r);
                case ReadRoute.RawRecords: return CsvDocument.ReadRecords(new StringReader(text), reader).ToArray();
#if NET8_0_OR_GREATER
                case ReadRoute.ReaderAsyncPath: using (var r = await CsvDocument.OpenReaderAsync(path, reader)) return Rows(r);
                case ReadRoute.ParallelTextRows:
                    return CsvDocument.ReadTextRowsAsParallel<string[]>(text, _ => RecordValues, reader,
                        new ParallelRowMappingOptions { MaxDegreeOfParallelism = 3, BatchSize = 7 }).ToArray();
                case ReadRoute.ParallelAggregate:
                    return CsvDocument.AggregateTextRowsAsParallel(text, static () => new List<string[]>(),
                        _ => (ref List<string[]> rows, CsvRecord row) => rows.Add(RecordValues(row)),
                        static (left, right) => { left.AddRange(right); return left; }, reader,
                        new ParallelRowMappingOptions { MaxDegreeOfParallelism = 1, BatchSize = 7 }).ToArray();
                case ReadRoute.FieldSpansText: return Collect(collect => CsvDocument.ReadFieldSpansFromText(text, collect, reader));
                case ReadRoute.FieldSpansReader: return Collect(collect => CsvDocument.ReadFieldSpans(new StringReader(text), collect, reader));
#else
                // These routes are .NET 8+ only; net472 checks the raw record path instead.
                case ReadRoute.ReaderAsyncPath:
                case ReadRoute.ParallelTextRows:
                case ReadRoute.ParallelAggregate:
                case ReadRoute.FieldSpansText:
                case ReadRoute.FieldSpansReader: return CsvDocument.ReadRecords(new StringReader(text), reader).ToArray();
#endif
                default: throw new ArgumentOutOfRangeException(nameof(route));
            }
        } finally { File.Delete(path); }
    }

#if NET8_0_OR_GREATER
    private static string[] RecordValues(CsvRecord row) {
        var values = new string[row.FieldCount];
        for (int i = 0; i < values.Length; i++) values[i] = row.GetString(i);
        return values;
    }
#endif

    private static string[][] Rows(CsvDocument document) =>
        document.AsEnumerable().Select(row => document.Header.Select(name => row.AsString(name) ?? string.Empty).ToArray()).ToArray();

    private static string[][] Rows(CsvDataReader reader) {
        var rows = new List<string[]>();
        while (reader.Read()) rows.Add(Enumerable.Range(0, reader.FieldCount).Select(reader.GetString).ToArray());
        return rows.ToArray();
    }

#if NET8_0_OR_GREATER
    private static string[][] Collect(Action<CsvFieldSpanAction> read) {
        var rows = new List<List<string>>();
        read((record, field, value) => {
            if (field == 0) rows.Add(new List<string>());
            rows[record].Add(value.ToString());
        });
        return rows.Select(row => row.ToArray()).ToArray();
    }
#endif
}
