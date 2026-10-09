#nullable enable
#if NET8_0_OR_GREATER
namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
    internal sealed partial class IncrementalRecords
    {
        private readonly CsvIncrementalFieldValues _borrowedFields = new();

        private bool CanBorrowFields => _delimiter.Length == 1 && !_options.InternStrings &&
            !_options.NormalizeQuotes && _options.MaxFieldLength is null;

        // Complete unquoted records can borrow the transport buffer until the next advance.
        // Buffer boundaries, comments, quotes and field transformations use the canonical parser.
        private bool TryReadBorrowedRecord(out bool emitted)
        {
            emitted = false;
            int start = _offset;
            if (_buffer[start] == _options.CommentCharacter) return false;
            int relativeEnd = _buffer.AsSpan(start, _length - start).IndexOfAny('\r', '\n', '"');
            if (relativeEnd < 0) return false;
            int end = start + relativeEnd;
            char ending = _buffer[end];
            if (ending == '"' || (ending == '\r' && end + 1 == _length)) return false;

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
    }
}
#endif
