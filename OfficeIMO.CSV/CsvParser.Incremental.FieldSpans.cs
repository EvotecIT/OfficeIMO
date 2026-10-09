#nullable enable
#if NET8_0_OR_GREATER
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
    internal sealed partial class IncrementalRecords
    {
        private readonly CsvIncrementalFieldValues _borrowedFields = new();
        // Only a transport optimization: avoid copying long tails that may still need canonical parsing.
        private const int MaximumBorrowedTailLength = 1024;

        private bool CanBorrowFields => _delimiter.Length == 1 && !_options.InternStrings &&
            !_options.NormalizeQuotes && _options.MaxFieldLength is null;

        // Complete unquoted records borrow the transport buffer until the next advance.
        // Short incomplete tails may append once; comments, quotes and transformations remain canonical.
        private bool TryReadBorrowedRecord(out bool emitted, out bool canRefill)
        {
            emitted = canRefill = false;
            int start = _offset;
            if (_buffer[start] == _options.CommentCharacter) return false;
            int relativeEnd = _buffer.AsSpan(start, _length - start).IndexOfAny('\r', '\n', '"');
            if (relativeEnd < 0)
            {
                canRefill = _length - start < _buffer.Length && _length - start <= MaximumBorrowedTailLength;
                return false;
            }
            int end = start + relativeEnd;
            char ending = _buffer[end];
            if (ending == '"') return false;
            if (ending == '\r' && end + 1 == _length)
            {
                canRefill = _length - start < _buffer.Length && _length - start <= MaximumBorrowedTailLength;
                return false;
            }

            _borrowedFields.SetFields(_buffer, start, relativeEnd, _delimiter[0], _options.TrimWhitespace);
            _fields.Clear();
            _offset = end + 1;
            if (ending == '\r' && _buffer[_offset] == '\n') _offset++;
            StartLine = EndLine = _line++;
            if (!_options.AllowEmptyLines && _borrowedFields.Count == 1 && _borrowedFields.GetSpan(0).IsEmpty)
                return true;

            Current = new CsvParsedRecord(_borrowedFields, startsWithCommentCharacter: false);
            ReportProgress(_options, ++_emitted, EndLine);
            emitted = true;
            return true;
        }

        /// <summary>
        /// Compacts an eligible unconsumed tail and appends once without advancing its physical line.
        /// The caller retries the same borrowed parser; quotes, EOF and incomplete rows remain canonical.
        /// Earlier borrowed fields have already been cleared by the current advance.
        /// </summary>
        private async ValueTask<bool> RefillBorrowedTailAsync(bool asynchronous, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            ThrowIfCancellationRequested(_options);
            int remaining = _length - _offset;
            if (_offset > 0) _buffer.AsSpan(_offset, remaining).CopyTo(_buffer.AsSpan());
            _offset = 0;
            _length = remaining;
            int read = asynchronous
                ? await _reader.ReadAsync(_buffer.AsMemory(remaining), token).ConfigureAwait(false)
                : _reader.Read(_buffer, remaining, _buffer.Length - remaining);
            _length += read;
            return read > 0;
        }
    }
}
#endif
