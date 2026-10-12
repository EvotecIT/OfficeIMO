#nullable enable
#if NET8_0_OR_GREATER
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
    internal sealed partial class IncrementalRecords
    {
        // These modes have observable serial validation or callback effects. Keep their
        // entire capture path canonical rather than moving any of those effects to workers.
        internal bool CanCaptureRawRecords => CanBorrowFields && _pending.Count == 0 &&
            _delimiter[0] is not '\r' and not '\n' and not '"' &&
            _options.ColumnCountMismatchPolicy != CsvColumnCountMismatchPolicy.Strict &&
            _options.ProgressCallback is null && !_options.CollectParseErrors &&
            _options.ParseErrorAction == CsvParseErrorAction.Throw &&
            _options.MaxQuotedFieldLength is null;

        /// <summary>
        /// Copies a run of complete unquoted records without splitting its fields. Any complex
        /// record stays unconsumed for the same parser's next canonical advance.
        /// </summary>
        // Internal callers await each result once. Reuse suspended refill state instead of
        // allocating another state-machine box for every transport read in a long scan.
        [System.Runtime.CompilerServices.AsyncMethodBuilder(typeof(System.Runtime.CompilerServices.PoolingAsyncValueTaskMethodBuilder<>))]
        internal async ValueTask<(int Count, bool ReachedEnd)> CaptureRawRecordRunAsync(
            CsvRecordTextBatch batch, CancellationToken token)
        {
            if (_failed) throw new InvalidOperationException("The CSV reader cannot continue after an interrupted record.");
            try
            {
                token.ThrowIfCancellationRequested();
                ThrowIfCancellationRequested(_options);
                // No borrowed current row may survive compaction or a transport refill.
                _borrowedFields.Clear();
                Current = default;
                if (!await EnsureBufferAsync(asynchronous: true, token).ConfigureAwait(false))
                    return (0, true);
                int count = CaptureRawRecordRun(batch, token, out bool canRefill);
                if (count == 0 && canRefill &&
                    await RefillBorrowedTailAsync(asynchronous: true, token).ConfigureAwait(false))
                    count = CaptureRawRecordRun(batch, token, out _);
                return (count, false);
            }
            catch { _failed = true; throw; }
        }

        private int CaptureRawRecordRun(CsvRecordTextBatch batch, CancellationToken token, out bool canRefill)
        {
            canRefill = false;
            int firstRow = batch.Count;
            int blockStart = _offset;
            int textOffset = batch.RawTextOffset;
            int lastStart = 0, lastLength = 0, lastLine = 0;
            batch.ConfigureRawRecords(_delimiter[0], _options.TrimWhitespace, _options.NullValue);
            while (_offset < _length && batch.Count < batch.RowCapacity)
            {
                token.ThrowIfCancellationRequested();
                ThrowIfCancellationRequested(_options);
                int start = _offset;
                if (_buffer[start] == _options.CommentCharacter) break;
                int relativeEnd = _buffer.AsSpan(start, _length - start).IndexOfAny('\r', '\n', '"');
                if (relativeEnd < 0)
                {
                    canRefill = batch.Count == firstRow && _length - start < _buffer.Length &&
                        _length - start <= MaximumBorrowedTailLength;
                    break;
                }
                int end = start + relativeEnd;
                char ending = _buffer[end];
                if (ending == '"') break;
                if (ending == '\r' && end + 1 == _length)
                {
                    canRefill = batch.Count == firstRow && _length - start < _buffer.Length &&
                        _length - start <= MaximumBorrowedTailLength;
                    break;
                }
                int next = end + 1;
                if (ending == '\r' && _buffer[next] == '\n') next++;
                bool empty = !_options.AllowEmptyLines && IsEmptyUnquotedRecord(start, relativeEnd);
                // Skipped leading lines need no detached text. Once a run has begun, its
                // separators and skipped lines are included in the single contiguous copy.
                if (empty && batch.Count == firstRow)
                {
                    _offset = blockStart = next;
                    _line++;
                    continue;
                }
                if (next - blockStart > batch.RemainingRawTextLength) break;
                _offset = next;
                int line = _line++;
                if (empty) continue;
                batch.CaptureRawRecord(textOffset + start - blockStart, relativeEnd);
                lastStart = start;
                lastLength = relativeEnd;
                lastLine = line;
            }
            int captured = batch.Count - firstRow;
            if (captured > 0)
            {
                batch.AppendRawText(_buffer, blockStart, _offset - blockStart);
                // Keep the ordinary reader positioned on the last captured source row. Only
                // that row needs producer field ranges; all batch rows are split by workers.
                _borrowedFields.SetFields(_buffer, lastStart, lastLength, _delimiter[0], _options.TrimWhitespace);
                _fields.Clear();
                Current = new CsvParsedRecord(_borrowedFields, startsWithCommentCharacter: false);
                StartLine = EndLine = lastLine;
                _emitted += captured;
            }
            return captured;
        }

        private bool IsEmptyUnquotedRecord(int start, int length)
        {
            if (length == 0) return true;
            if (!_options.TrimWhitespace) return false;
            // A whitespace delimiter still separates multiple empty fields, which is an
            // emitted record. Match the canonical one-field, trimmed-empty rule exactly.
            ReadOnlySpan<char> value = _buffer.AsSpan(start, length);
            if (value.IndexOf(_delimiter[0]) >= 0) return false;
            foreach (char character in value)
                if (!char.IsWhiteSpace(character)) return false;
            return true;
        }
    }
}
#endif
