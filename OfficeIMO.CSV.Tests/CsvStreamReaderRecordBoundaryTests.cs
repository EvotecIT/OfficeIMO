using System.Text;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

/// <summary>
/// Stream readers return the same fields, physical lines, progress lines and parse-error lines for
/// a final record without a separator after earlier records, blank lines, short reads, trimming and
/// multiline quoted records, whichever decoding route the encoding selects.
/// </summary>
public sealed class CsvStreamReaderRecordBoundaryTests {
    public static TheoryData<string, string, bool, string, bool> Cases => new() {
        { "x,y\ra,b", "x|y@1/1 a|b@2/2", false, "", false },
        { "\ra,b", "a|b@2/2", false, "", false },
        { "\n\nab,cd", "ab|cd@3/3", false, "", false },
        { "\nbc", "bc@2/2", true, "", false },
        { "\n  bc ,d", "bc|d@2/2", true, "", false },
        { "bc,\"x\"\r\n \r\ra,,\"\"\nbca\r\nbc", "bc|x@1/1  |@2/2 a|@4/4 bca|@5/5 bc|@6/6", false, "", false },
        { "a,1\n\"q\r\nr\",2\n\"x\"y,3\nb,4", "a|1@1/1 q\r\nr|2@2/2 b|4@5/5", false, "4", false },
        { "\"q\nr\",1\n\"s\nt\",2\n\"u", "q\nr|1@1/1 s\nt|2@3/3", false, "5", false },
        { "\né,é", "é|é@2/2", false, "", false },
        { "#a,b\nc,d", "#a|b@1/1 c|d@2/2", false, "", false },
        { "a,b\n#x\r\nc,d", "a|b@1/1 c|d@3/3", false, "", true },
        { "a,b\n#x \"y\nz\"\nc,d", "a|b@1/1 c|d@4/4", false, "", true },
        { "#x \"y\na,b\nq\"\nc,d", "a|b@2/2 c|d@4/4", false, "3", true }
    };

    [Theory]
    [MemberData(nameof(Cases))]
    public void FinalRecordsKeepTheirFieldsAndLines(string input, string expectedRows, bool trim, string expectedErrorLines, bool skipComments) {
        // Default UTF-8 selects the byte reader. On .NET 8 and later, Latin-1 and a non-U+FFFD UTF-8
        // fallback select the decoded text reader; older targets read those through the document parser.
        var encodings = new List<(Encoding? Encoding, byte[] Bytes)> { (null, Encoding.UTF8.GetBytes(input)) };
#if NETCOREAPP
        encodings.Add((Encoding.GetEncoding("utf-8", EncoderFallback.ReplacementFallback, new DecoderReplacementFallback("?")), Encoding.UTF8.GetBytes(input)));
        if (input.All(character => character < 256)) encodings.Add((Encoding.Latin1, Encoding.Latin1.GetBytes(input)));
#endif

        foreach ((Encoding? encoding, byte[] bytes) in encodings) {
            foreach (int chunk in new[] { 0, 1, 3 }) {
                var errors = new List<CsvParseError>();
                int progressLine = 0;
                var options = new CsvReaderOptions {
                    HasHeaderRow = false,
                    DetectDelimiter = false,
                    Encoding = encoding,
                    TrimWhitespace = trim,
                    SkipCommentRows = skipComments,
                    QuoteParsingMode = CsvQuoteParsingMode.Strict,
                    ParseErrorAction = CsvParseErrorAction.SkipRow,
                    CollectParseErrors = true,
                    ParseErrors = errors,
                    ProgressReportInterval = 1,
                    ProgressCallback = progress => progressLine = progress.LineNumber
                };
                var rows = new List<string>();
                using (CsvDataReader reader = CsvDocument.OpenReader(new ShortReadStream(bytes, chunk), options)) {
                    var position = (ICsvDataReaderPositionMetadata)reader;
                    while (reader.Read()) {
                        var fields = new string[reader.FieldCount];
                        for (int column = 0; column < fields.Length; column++) fields[column] = reader.GetString(column);
                        rows.Add($"{string.Join("|", fields)}@{position.PhysicalLineNumber}/{progressLine}");
                    }
                }

                string route = $"{encoding?.WebName ?? "default"} {encoding?.DecoderFallback.GetType().Name} chunk {chunk}";
                Assert.True(expectedRows == string.Join(" ", rows), $"{route}: {string.Join(" ", rows).Replace("\r", "\\r").Replace("\n", "\\n")}");
                Assert.True(expectedErrorLines == string.Join(" ", errors.Select(error => error.LineNumber)),
                    $"{route}: errors at {string.Join(" ", errors.Select(error => error.LineNumber))}");
            }
        }
    }

    /// <summary>Seekable input whose reads return at most a few bytes.</summary>
    private sealed class ShortReadStream : MemoryStream {
        private readonly int _chunk;

        internal ShortReadStream(byte[] bytes, int chunk) : base(bytes, writable: false) => _chunk = chunk;

        public override int Read(byte[] buffer, int offset, int count) =>
            base.Read(buffer, offset, _chunk > 0 ? Math.Min(count, _chunk) : count);

#if NETCOREAPP
        public override int Read(Span<byte> buffer) =>
            base.Read(_chunk > 0 && buffer.Length > _chunk ? buffer.Slice(0, _chunk) : buffer);
#endif
    }
}
