using System.Text;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

/// <summary>
/// The UTF-8 stream reader parses quoted records from bytes. Its rows, physical lines, progress and
/// parse errors must follow the RFC 4180 rules, the lenient fallback and the error policy for any
/// mix of quoting, line separators and short stream reads.
/// </summary>
public sealed class CsvUtf8QuotedRecordParityTests {
    private static readonly string[] Pieces = {
        "a", "bc", "é", " ", ",", ",", "\"", "\"\"", "\"x\"", "\r", "\n", "\r\n", "\"q,\r\nr\""
    };

    [Theory]
    [InlineData(CsvQuoteParsingMode.Lenient, CsvParseErrorAction.Throw)]
    [InlineData(CsvQuoteParsingMode.Strict, CsvParseErrorAction.Throw)]
    [InlineData(CsvQuoteParsingMode.Strict, CsvParseErrorAction.SkipRow)]
    public void StreamReaderFollowsQuotingRulesForRandomInput(CsvQuoteParsingMode quoting, CsvParseErrorAction errors) {
        var random = new Random(20261110 + (int)quoting * 7 + (int)errors);
        for (int sample = 0; sample < 1000; sample++) {
            string input = sample % 2 == 0 ? RandomPieces(random) : RandomRecords(random);
            bool allowEmpty = random.Next(2) == 0;
            string expected = new QuotingModel(input, quoting == CsvQuoteParsingMode.Strict,
                errors == CsvParseErrorAction.SkipRow, allowEmpty).Run();
            string actual = Observe(input, quoting, errors, allowEmpty, random.Next(0, 8));
            if (expected != actual) {
                Assert.Fail($"Input {Escape(input)} (allowEmpty {allowEmpty}):\nexpected\n{Escape(expected)}\nactual\n{Escape(actual)}");
            }
        }
    }

    [Fact]
    public void QuotedRecordsCrossingBufferRefillsKeepTheirText() {
        var text = new StringBuilder();
        for (int row = 0; row < 12_000; row++) {
            text.Append(row).Append(",\"välue ").Append(row).Append(" \"\"q\"\"\r\nnext\",").Append(new string('x', row % 50)).Append('\n');
        }
        string input = text.ToString();
        Assert.Equal(
            new QuotingModel(input, strict: true, skip: false, allowEmpty: false).Run(),
            Observe(input, CsvQuoteParsingMode.Strict, CsvParseErrorAction.Throw, allowEmpty: false, chunk: 0));
    }

    private static string RandomPieces(Random random) {
        var text = new StringBuilder();
        int pieces = random.Next(1, 120);
        for (int piece = 0; piece < pieces; piece++) text.Append(Pieces[random.Next(Pieces.Length)]);
        return text.ToString();
    }

    // Mostly well-formed records long enough for whole 64-byte blocks, with occasional damage.
    private static string RandomRecords(Random random) {
        string[] separators = { "\n", "\r\n", "\r" };
        var text = new StringBuilder();
        int records = random.Next(1, 8);
        for (int record = 0; record < records; record++) {
            int fields = random.Next(1, 12);
            for (int field = 0; field < fields; field++) {
                if (field > 0) text.Append(',');
                int length = random.Next(0, 40);
                bool quoted = random.Next(3) > 0;
                if (quoted) text.Append('"');
                for (int index = 0; index < length; index++) {
                    int kind = random.Next(30);
                    text.Append(kind switch {
                        0 when quoted => "\"\"",
                        1 when quoted => ",",
                        2 when quoted => separators[random.Next(separators.Length)],
                        3 => "é",
                        _ => ((char)('a' + kind % 26)).ToString()
                    });
                }
                if (quoted) text.Append('"');
                if (random.Next(60) == 0) text.Append('"');
            }
            text.Append(separators[random.Next(separators.Length)]);
        }
        if (random.Next(4) == 0) text.Length -= random.Next(1, Math.Min(4, text.Length) + 1);
        return text.ToString();
    }

    private static string Escape(string value) =>
        value.Replace("\r", "\\r").Replace("\n", "\\n\n");

    private static string Observe(
        string input,
        CsvQuoteParsingMode quoting,
        CsvParseErrorAction errors,
        bool allowEmpty,
        int chunk) {
        var result = new StringBuilder();
        var parseErrors = new List<CsvParseError>();
        int progressLine = 0;
        var options = new CsvReaderOptions {
            HasHeaderRow = false,
            DetectDelimiter = false,
            QuoteParsingMode = quoting,
            ParseErrorAction = errors,
            CollectParseErrors = errors == CsvParseErrorAction.SkipRow,
            ParseErrors = parseErrors,
            AllowEmptyLines = allowEmpty,
            ProgressReportInterval = 1,
            ProgressCallback = progress => progressLine = progress.LineNumber
        };
        try {
            using var stream = new ChunkedStream(Encoding.UTF8.GetBytes(input), chunk);
            using CsvDataReader reader = CsvDocument.OpenReader(stream, options);
            var position = (ICsvDataReaderPositionMetadata)reader;
            while (reader.Read()) {
                result.Append("row ").Append(position.RecordNumber).Append(' ')
                    .Append(position.PhysicalLineNumber).Append('-').Append(position.PhysicalEndLineNumber)
                    .Append(" progress ").Append(progressLine).Append(':');
                for (int column = 0; column < reader.FieldCount; column++) {
                    result.Append(" [").Append(reader.GetString(column)).Append(']');
                }
                result.Append('\n');
            }
        } catch (CsvParseException exception) {
            result.Append("error ").Append(exception.LineNumber).Append(' ').Append(exception.Message).Append('\n');
        }
        foreach (CsvParseError error in parseErrors) {
            result.Append("collected ").Append(error.LineNumber).Append(' ').Append(error.Message).Append('\n');
        }
        return result.ToString();
    }

    /// <summary>
    /// Reference reading of the documented quoting rules: an RFC 4180 record, else in lenient mode a
    /// flexible record in which any quote opens a quoted section and a quote right after a closing
    /// quote is literal. Lines are physical lines; CR, LF and CRLF each end one line.
    /// </summary>
    private sealed class QuotingModel {
        private readonly string _input;
        private readonly bool _strict;
        private readonly bool _skip;
        private readonly bool _allowEmpty;
        private readonly StringBuilder _result = new();
        private readonly StringBuilder _collected = new();
        private int _position;
        private int _line = 1;
        private int _records;
        private int _fieldCount = -1;

        internal QuotingModel(string input, bool strict, bool skip, bool allowEmpty) {
            _input = input;
            _strict = strict;
            _skip = skip;
            _allowEmpty = allowEmpty;
        }

        internal string Run() {
            while (_position < _input.Length) {
                int start = _position;
                int startLine = _line;
                var fields = new List<string>();
                int end = ParseStandard(start, fields, out bool unterminated, out int invalidAt);
                if (invalidAt >= 0 && !_strict) {
                    fields.Clear();
                    end = ParseFlexible(start, fields, out unterminated);
                    invalidAt = -1;
                }

                if (unterminated) {
                    // The rest of the input belongs to the open quote; its line is the last line read.
                    Error(startLine + Separators(start, _input.Length) - (EndsWithSeparator() ? 1 : 0), "Unterminated quoted field.");
                    break;
                }

                if (invalidAt >= 0) {
                    if (!Error(startLine + Separators(start, invalidAt), "Invalid quoted field.")) break;
                    int next = _input.IndexOfAny(new[] { '\r', '\n' }, invalidAt);
                    Advance(start, next < 0 ? _input.Length : next);
                    continue;
                }

                Advance(start, end);
                if (fields.Count == 1 && fields[0].Length == 0 && !_allowEmpty) continue;
                if (_fieldCount < 0) _fieldCount = fields.Count;
                _records++;
                _result.Append("row ").Append(_records).Append(' ').Append(startLine).Append('-')
                    .Append(startLine + Separators(start, end)).Append(" progress ").Append(startLine).Append(':');
                for (int column = 0; column < _fieldCount; column++) {
                    _result.Append(" [").Append(column < fields.Count ? fields[column] : string.Empty).Append(']');
                }
                _result.Append('\n');
            }

            return _result.Append(_collected).ToString();
        }

        private bool Error(int line, string message) {
            string text = $"Line {line}: {message}";
            if (!_skip) {
                _result.Append("error ").Append(line).Append(' ').Append(text).Append('\n');
                return false;
            }
            _collected.Append("collected ").Append(line).Append(' ').Append(text).Append('\n');
            return true;
        }

        // Moves past a record ending at a separator index (or the input end) and counts its lines.
        private void Advance(int start, int end) {
            int next = end;
            if (next < _input.Length) {
                next += _input[next] == '\r' && next + 1 < _input.Length && _input[next + 1] == '\n' ? 2 : 1;
            }
            _line += Separators(start, next);
            _position = next;
        }

        private int Separators(int start, int end) {
            int count = 0;
            for (int index = start; index < end; index++) {
                if (_input[index] == '\n' || (_input[index] == '\r' && (index + 1 >= end || _input[index + 1] != '\n'))) count++;
            }
            return count;
        }

        private bool EndsWithSeparator() => _input.Length > 0 && _input[_input.Length - 1] is '\r' or '\n';

        private int ParseStandard(int position, List<string> fields, out bool unterminated, out int invalidAt) {
            unterminated = false;
            invalidAt = -1;
            int fieldStart = position;
            while (true) {
                if (position == _input.Length || _input[position] is '\r' or '\n') {
                    fields.Add(_input.Substring(fieldStart, position - fieldStart));
                    return position;
                }
                char value = _input[position];
                if (value == ',') {
                    fields.Add(_input.Substring(fieldStart, position - fieldStart));
                    fieldStart = ++position;
                    continue;
                }
                if (value != '"') {
                    position++;
                    continue;
                }
                if (position != fieldStart) {
                    invalidAt = position;
                    return -1;
                }
                var quoted = new StringBuilder();
                position++;
                while (true) {
                    if (position == _input.Length) {
                        unterminated = true;
                        return -1;
                    }
                    if (_input[position] != '"') {
                        quoted.Append(_input[position++]);
                    } else if (position + 1 < _input.Length && _input[position + 1] == '"') {
                        quoted.Append('"');
                        position += 2;
                    } else {
                        position++;
                        break;
                    }
                }
                fields.Add(quoted.ToString());
                if (position == _input.Length || _input[position] is '\r' or '\n') return position;
                if (_input[position] != ',') {
                    invalidAt = position;
                    return -1;
                }
                fieldStart = ++position;
            }
        }

        private int ParseFlexible(int position, List<string> fields, out bool unterminated) {
            var field = new StringBuilder();
            bool inQuotes = false;
            bool afterClosingQuote = false;
            for (; position < _input.Length; position++) {
                char value = _input[position];
                if (inQuotes) {
                    if (value != '"') {
                        field.Append(value);
                    } else if (position + 1 < _input.Length && _input[position + 1] == '"') {
                        field.Append('"');
                        position++;
                    } else {
                        inQuotes = false;
                        afterClosingQuote = true;
                    }
                    continue;
                }
                if (value == '"') {
                    if (afterClosingQuote) field.Append('"');
                    inQuotes = !afterClosingQuote;
                    afterClosingQuote = false;
                    continue;
                }
                if (value is '\r' or '\n') break;
                if (value == ',') {
                    fields.Add(field.ToString());
                    field.Clear();
                } else {
                    field.Append(value);
                }
                afterClosingQuote = false;
            }
            unterminated = inQuotes;
            fields.Add(field.ToString());
            return position;
        }
    }

    /// <summary>Seekable input that returns at most a few bytes per read to force refills mid-record.</summary>
    private sealed class ChunkedStream : MemoryStream {
        private readonly int _chunk;

        internal ChunkedStream(byte[] bytes, int chunk) : base(bytes, writable: false) => _chunk = chunk;

        public override int Read(byte[] buffer, int offset, int count) =>
            base.Read(buffer, offset, _chunk > 0 ? Math.Min(count, _chunk) : count);

        public override int Read(Span<byte> buffer) =>
            base.Read(_chunk > 0 && buffer.Length > _chunk ? buffer.Slice(0, _chunk) : buffer);
    }
}
