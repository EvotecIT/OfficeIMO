using System.Text;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

/// <summary>
/// Field length limits and multi-character delimiters read from a byte stream give the same records,
/// parse-error lines, messages and skip behaviour as the canonical record parser over the same text.
/// </summary>
public sealed class CsvFieldLimitAndTextDelimiterParityTests {
    private static readonly string[] Pieces = {
        "a", "bcd", "é", "日本", " ", "\t", "\u00a0", "\"", "\"\"", "\"x\"", "\r", "\n", "\r\n", "#"
    };

    [Theory]
    [InlineData(",", CsvQuoteParsingMode.Lenient, CsvParseErrorAction.SkipRow)]
    [InlineData(",", CsvQuoteParsingMode.Strict, CsvParseErrorAction.SkipRow)]
    [InlineData(",", CsvQuoteParsingMode.Lenient, CsvParseErrorAction.Throw)]
    [InlineData("||", CsvQuoteParsingMode.Lenient, CsvParseErrorAction.SkipRow)]
    [InlineData("||", CsvQuoteParsingMode.Strict, CsvParseErrorAction.SkipRow)]
    [InlineData("||", CsvQuoteParsingMode.Strict, CsvParseErrorAction.Throw)]
    [InlineData("<->", CsvQuoteParsingMode.Lenient, CsvParseErrorAction.SkipRow)]
    [InlineData(" ;", CsvQuoteParsingMode.Strict, CsvParseErrorAction.SkipRow)]
    public void StreamedRecordsMatchCanonicalParser(string delimiter, CsvQuoteParsingMode quoting, CsvParseErrorAction errors) {
        var random = new Random(delimiter.Length * 1000 + (int)quoting * 10 + (int)errors);
        string[] pieces = Pieces.Concat(new[] { delimiter, delimiter, delimiter.Substring(0, 1), "\"" + delimiter + "\"" }).ToArray();
        int compared = 0;
        for (int sample = 0; sample < 1500; sample++) {
            var text = new StringBuilder();
            int count = random.Next(1, 60);
            for (int piece = 0; piece < count; piece++) text.Append(pieces[random.Next(pieces.Length)]);
            string input = text.ToString();
            CsvReaderOptions options = CreateOptions(delimiter, quoting, errors, random);
            string expected = ObserveCanonical(input, options);
            // Only a one-character delimiter on .NET 8 and later has a decoded text reader to compare.
#if NETCOREAPP
            bool decoded = delimiter.Length == 1 && random.Next(3) == 0;
#else
            bool decoded = false;
#endif
            int chunk = random.Next(0, 6);
            // With a one-character delimiter, the streaming readers report a strict-mode invalid quote on
            // its own line where the canonical parser first joins the next line; those inputs are not
            // about field limits, so only inputs both routes read alike without limits are compared.
            if (delimiter.Length == 1 && quoting == CsvQuoteParsingMode.Strict) {
                CsvReaderOptions unlimited = options.Clone();
                unlimited.MaxFieldLength = null;
                unlimited.MaxQuotedFieldLength = null;
                if (ObserveCanonical(input, unlimited) != ObserveStream(input, unlimited, chunk, decoded)) continue;
            }
            compared++;
            string actual = ObserveStream(input, options, chunk, decoded);
            if (expected != actual) {
                Assert.Fail($"Input {Escape(input)} (max {options.MaxFieldLength}, quoted max {options.MaxQuotedFieldLength}, " +
                    $"decoded {decoded}, trim {options.TrimWhitespace}, empty {options.AllowEmptyLines}, comments {options.SkipCommentRows}):\n" +
                    $"expected\n{Escape(expected)}\nactual\n{Escape(actual)}");
            }
        }
        Assert.True(compared > 1000, $"Only {compared} inputs were compared.");
    }

    [Theory]
    [InlineData(",")]
    [InlineData("||")]
    public void LongRecordsCrossingBufferRefillsMatchCanonicalParser(string delimiter) {
        var text = new StringBuilder();
        for (int row = 0; row < 9_000; row++) {
            text.Append(row).Append(delimiter).Append("\"välue ").Append(row).Append(" \"\"q\"\"\r\nnext\"")
                .Append(delimiter).Append(new string('x', row % 70)).Append('\n');
        }
        string input = text.ToString();
        var options = new CsvReaderOptions {
            HasHeaderRow = false, DetectDelimiter = false, DelimiterText = delimiter, MaxFieldLength = 60, MaxQuotedFieldLength = 25,
            ParseErrorAction = CsvParseErrorAction.SkipRow, CollectParseErrors = true, ParseErrors = new List<CsvParseError>()
        };
        Assert.Equal(ObserveCanonical(input, options), ObserveStream(input, options, chunk: 0, decoded: false));
#if NETCOREAPP
        Assert.Equal(ObserveCanonical(input, options), ObserveStream(input, options, chunk: 0, decoded: true));
#endif
    }

    private static CsvReaderOptions CreateOptions(string delimiter, CsvQuoteParsingMode quoting, CsvParseErrorAction errors, Random random) =>
        new() {
            HasHeaderRow = false,
            DetectDelimiter = false,
            DelimiterText = delimiter,
            QuoteParsingMode = quoting,
            ParseErrorAction = errors,
            CollectParseErrors = errors == CsvParseErrorAction.SkipRow,
            AllowEmptyLines = random.Next(2) == 0,
            TrimWhitespace = random.Next(2) == 0,
            SkipCommentRows = random.Next(4) == 0,
            MaxFieldLength = random.Next(3) == 0 ? null : random.Next(0, 7),
            MaxQuotedFieldLength = random.Next(3) == 0 ? null : random.Next(0, 5)
        };

    // The canonical parser keeps every field of every record; a headerless data reader exposes the
    // first record's width and reads a missing field as empty text.
    private static string ObserveCanonical(string input, CsvReaderOptions options) {
        options = options.Clone();
        var errors = new List<CsvParseError>();
        options.ParseErrors = errors;
        var result = new StringBuilder();
        int width = -1;
        try {
            foreach (IReadOnlyList<string> record in CsvParser.ParseReusable(new StringReader(input), options.Parsing)) {
                if (width < 0) width = record.Count;
                for (int column = 0; column < width; column++) {
                    result.Append(column < record.Count ? record[column] : "").Append('|');
                }
                result.Append('\n');
            }
        } catch (CsvParseException exception) {
            result.Append("error ").Append(exception.LineNumber).Append(' ').Append(exception.Message).Append('\n');
        }
        return Describe(result, errors);
    }

    // A UTF-8 decoder that replaces malformed input with "?" selects the decoded text reader.
    private static string ObserveStream(string input, CsvReaderOptions options, int chunk, bool decoded) {
        options = options.Clone();
        if (decoded) options.Encoding = Encoding.GetEncoding("utf-8", EncoderFallback.ReplacementFallback, new DecoderReplacementFallback("?"));
        var errors = new List<CsvParseError>();
        options.ParseErrors = errors;
        var result = new StringBuilder();
        try {
            using var stream = new ChunkedStream(Encoding.UTF8.GetBytes(input), chunk);
            using CsvDataReader reader = CsvDocument.OpenReader(stream, options);
            while (reader.Read()) {
                for (int column = 0; column < reader.FieldCount; column++) {
                    result.Append(reader.GetString(column)).Append('|');
                }
                result.Append('\n');
            }
        } catch (CsvParseException exception) {
            result.Append("error ").Append(exception.LineNumber).Append(' ').Append(exception.Message).Append('\n');
        }
        return Describe(result, errors);
    }

    private static string Describe(StringBuilder result, List<CsvParseError> errors) {
        foreach (CsvParseError error in errors) result.Append("skipped ").Append(error.LineNumber).Append(' ').Append(error.Message).Append('\n');
        return result.ToString();
    }

    private static string Escape(string value) => value.Replace("\r", "\\r").Replace("\n", "\\n\n");

    /// <summary>Seekable input whose reads return at most a few bytes.</summary>
    private sealed class ChunkedStream : MemoryStream {
        private readonly int _chunk;

        internal ChunkedStream(byte[] bytes, int chunk) : base(bytes, writable: false) => _chunk = chunk;

        public override int Read(byte[] buffer, int offset, int count) =>
            base.Read(buffer, offset, _chunk > 0 ? Math.Min(count, _chunk) : count);

#if NETCOREAPP
        public override int Read(Span<byte> buffer) =>
            base.Read(_chunk > 0 && buffer.Length > _chunk ? buffer.Slice(0, _chunk) : buffer);
#endif
    }
}
