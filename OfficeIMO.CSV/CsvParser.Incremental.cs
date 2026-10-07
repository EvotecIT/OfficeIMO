#nullable enable
#if NET8_0_OR_GREATER
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
    // The transport is asynchronous; field splitting and validation use the same grammar as Parse.
    internal sealed class IncrementalRecords : IDisposable
    {
        private readonly TextReader _reader;
        private readonly CsvLoadOptions _options;
        private readonly char[] _buffer = new char[4096];
        private readonly Queue<CsvLine> _pending = new();
        private readonly Dictionary<string, string>? _cache;
        private readonly string _delimiter;
        private readonly List<string> _fields = new();
        private StringBuilder? _logicalRecordBuffer;
        private StringBuilder? _lineBuffer;
        private const int MaximumRetainedTextCapacity = 64 * 1024;
        private int _offset, _length, _line = 1, _emitted;
        private bool _failed;
        internal CsvParsedRecord Current { get; private set; }
        internal int StartLine { get; private set; }
        internal int EndLine { get; private set; }

        internal IncrementalRecords(TextReader reader, CsvLoadOptions options)
        {
            _reader = reader;
            _options = options;
            _cache = CreateStringCache(options);
            _delimiter = GetDelimiterText(options);
        }

        // Initialization retains schema samples; normal traversal borrows fields until the next advance.
        internal async ValueTask<bool> ReadAsync(bool asynchronous, CancellationToken token, bool reuseValues = false)
        {
            if (_failed) throw new InvalidOperationException("The CSV reader cannot continue after an interrupted record.");
            try
            {
                while (await ReadLineAsync(asynchronous, token).ConfigureAwait(false) is { } first)
                {
                    token.ThrowIfCancellationRequested();
                    ThrowIfCancellationRequested(_options);
                    StartLine = first.PhysicalLineNumber;
                    var delimiter = _delimiter;
                    var comment = IsRawCommentLine(first.Text, _options);
                    if (ShouldSkipCommentRecordBeforeParsing(comment, first.Text, _options, _emitted))
                    {
                        await SkipCommentAsync(first, delimiter, asynchronous, token).ConfigureAwait(false);
                        continue;
                    }

                    var quoted = first.Text.IndexOf('"') >= 0;
                    List<string> fields = _fields;
                    try
                    {
                        if (!quoted)
                        {
                            if (delimiter.Length == 1)
                                TrySplitUnquotedRecord(first.Text, delimiter[0], _options.TrimWhitespace, fields);
                            else
                            {
                                TrySplitUnquotedRecord(first.Text, delimiter, _options.TrimWhitespace, out string[] values);
                                fields.Clear();
                                fields.AddRange(values);
                            }
                            EndLine = first.PhysicalLineNumber;
                        }
                        else
                        {
                            var strict = _options.QuoteParsingMode == CsvQuoteParsingMode.Strict;
                            var record = first.Text;
                            EndLine = first.PhysicalLineNumber;
                            if (!TryParseIntoFields(record, delimiter, strict, EndLine))
                            {
                                var text = _logicalRecordBuffer ??= new StringBuilder();
                                text.Clear();
                                text.Append(record);
                                var separator = first.Separator;
                                var state = new QuotedRecordState();
                                UpdateState(record, delimiter, ref state);
                                while (state.InQuotes)
                                {
                                    var next = await ReadLineAsync(asynchronous, token).ConfigureAwait(false);
                                    if (next is null) throw new CsvParseException("Unterminated quoted field.", EndLine);
                                    text.Append(separator).Append(next.Value.Text);
                                    EndLine = next.Value.PhysicalLineNumber;
                                    separator = next.Value.Separator;
                                    UpdateState(next.Value.Text, delimiter, ref state);
                                }
                                string complete = text.ToString();
                                if (text.Capacity > MaximumRetainedTextCapacity) _logicalRecordBuffer = null;
                                if (!TryParseIntoFields(complete, delimiter, strict, EndLine))
                                    throw new CsvParseException("Unterminated quoted field.", EndLine);
                            }
                        }
                    }
                    catch (CsvParseException error) when (HandleParseError(_options, error, EndLine))
                    {
                        continue;
                    }

                    if (!ShouldEmitRecord(fields, _options.AllowEmptyLines) ||
                        ShouldSkipCommentRecord(comment, first.Text, _options, _emitted) ||
                        !TryPrepareParsedRecord(fields, _options, EndLine, quoted, _cache)) continue;
                    Current = new CsvParsedRecord(reuseValues ? fields : fields.ToArray(), comment);
                    ReportProgress(_options, ++_emitted, EndLine);
                    return true;
                }
                return false;
            }
            catch
            {
                _failed = true;
                throw;
            }
        }

        private bool TryParse(string text, string delimiter, bool strict, int line, out string[] fields) =>
            delimiter.Length == 1
                ? TryParseQuotedRecord(text, delimiter[0], _options.TrimWhitespace, strict, line, out fields)
                : TryParseQuotedRecord(text, delimiter, _options.TrimWhitespace, strict, line, out fields);

        private bool TryParseIntoFields(string text, string delimiter, bool strict, int line)
        {
            if (delimiter.Length == 1)
                return TryParseQuotedRecord(text, delimiter[0], _options.TrimWhitespace, strict, line, _fields);
            bool parsed = TryParse(text, delimiter, strict, line, out string[] values);
            _fields.Clear();
            if (parsed) _fields.AddRange(values);
            return parsed;
        }

        private void UpdateState(string text, string delimiter, ref QuotedRecordState state)
        {
            if (delimiter.Length == 1) UpdateQuotedRecordState(text, delimiter[0], _options.TrimWhitespace, ref state);
            else UpdateQuotedRecordState(text, delimiter, _options.TrimWhitespace, ref state);
        }

        private async ValueTask SkipCommentAsync(CsvLine first, string delimiter, bool asynchronous, CancellationToken token)
        {
            if (TryParse(first.Text, delimiter, false, first.PhysicalLineNumber, out _)) return;
            if (_options.MaxCommentContinuationLines <= 0)
                throw new ArgumentOutOfRangeException(nameof(_options.MaxCommentContinuationLines));
            var state = new QuotedRecordState();
            UpdateState(first.Text, delimiter, ref state);
            var continuations = new List<CsvLine>();
            while (continuations.Count < _options.MaxCommentContinuationLines)
            {
                var next = await ReadLineAsync(asynchronous, token).ConfigureAwait(false);
                if (next is null) break;
                continuations.Add(next.Value);
                UpdateState(next.Value.Text, delimiter, ref state);
                if (!state.InQuotes) return;
                if (!LooksLikeDelimitedRawComment(first.Text, delimiter) && LooksLikeDelimitedRawComment(next.Value.Text, delimiter)) break;
            }
            foreach (var line in continuations) _pending.Enqueue(line);
        }

        private async ValueTask<CsvLine?> ReadLineAsync(bool asynchronous, CancellationToken token)
        {
            if (_pending.Count > 0) return _pending.Dequeue();
            StringBuilder? text = null;
            while (true)
            {
                token.ThrowIfCancellationRequested();
                if (!await EnsureBufferAsync(asynchronous, token).ConfigureAwait(false))
                    return text is null ? null : new CsvLine(CompleteBufferedLine(text), string.Empty, _line++);
                int start = _offset;
                while (_offset < _length && _buffer[_offset] is not '\r' and not '\n') _offset++;
                if (_offset == _length)
                {
                    if (text is null)
                    {
                        text = _lineBuffer ??= new StringBuilder();
                        text.Clear();
                    }
                    text.Append(_buffer, start, _offset - start);
                    continue;
                }
                string value = text is null ? new string(_buffer, start, _offset - start)
                    : CompleteBufferedLine(text.Append(_buffer, start, _offset - start));
                char separator = _buffer[_offset++];
                string ending = separator == '\n' ? "\n" : "\r";
                if (separator == '\r' && await EnsureBufferAsync(asynchronous, token).ConfigureAwait(false) && _buffer[_offset] == '\n')
                {
                    _offset++;
                    ending = "\r\n";
                }
                return new CsvLine(value, ending, _line++);
            }
        }

        private string CompleteBufferedLine(StringBuilder text)
        {
            string value = text.ToString();
            if (text.Capacity > MaximumRetainedTextCapacity) _lineBuffer = null;
            return value;
        }

        private async ValueTask<bool> EnsureBufferAsync(bool asynchronous, CancellationToken token)
        {
            if (_offset < _length) return true;
            _offset = 0;
            _length = asynchronous
                ? await _reader.ReadAsync(_buffer.AsMemory(), token).ConfigureAwait(false)
                : _reader.Read(_buffer, 0, _buffer.Length);
            return _length > 0;
        }

        public void Dispose()
        {
            _fields.Clear();
            _logicalRecordBuffer = _lineBuffer = null;
            Current = default;
            _reader.Dispose();
        }
    }
}
#endif
