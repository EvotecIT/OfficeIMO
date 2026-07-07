#nullable enable

namespace OfficeIMO.CSV;

internal sealed partial class CsvLineReader
{
#if NET8_0_OR_GREATER
    private const int StandardQuotedFieldSpanCapacity = 64;
    private const int DirectUnescapeQuotedFieldMinimumLength = 32;

    private bool TryReadStandardQuotedFieldSpansOrLine<TVisitor>(
        char delimiter,
        bool trim,
        bool allowEmpty,
        bool emitFields,
        int recordIndex,
        ref TVisitor fieldVisitor,
        out int fieldCount,
        out bool isEmptyRecord,
        out string separator,
        out CsvLineReadResult readResult)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        fieldCount = 0;
        isEmptyRecord = false;
        separator = string.Empty;
        readResult = CsvLineReadResult.Line;

        var start = _position;
        Span<StandardCsvFieldSpan> fields = stackalloc StandardCsvFieldSpan[StandardQuotedFieldSpanCapacity];
        if (!TryParseStandardQuotedFieldSpans(start, delimiter, trim, fields, out var fieldCountValue, out var firstFieldLength, out var recordEnd))
        {
            return false;
        }

        var emit = emitFields && (allowEmpty || recordEnd > start);
        fieldCount = fieldCountValue;
        if (emit)
        {
            VisitStandardQuotedFieldSpans(fields.Slice(0, fieldCount), recordIndex, ref fieldVisitor);
        }

        isEmptyRecord = fieldCount == 1 && firstFieldLength == 0;
        _position = recordEnd;
        ConsumeLineSeparator(_buffer[recordEnd], out separator);
        readResult = CsvLineReadResult.UnquotedRecord;
        return true;
    }

    private bool TryReadStandardQuotedFieldSpansOrLineFromPrefix<TVisitor>(
        char delimiter,
        bool allowEmpty,
        bool emitFields,
        int recordIndex,
        ReadOnlySpan<int> delimiterIndexesBeforeQuote,
        int quoteIndex,
        ref TVisitor fieldVisitor,
        out int fieldCount,
        out bool isEmptyRecord,
        out string separator,
        out CsvLineReadResult readResult)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        fieldCount = 0;
        isEmptyRecord = false;
        separator = string.Empty;
        readResult = CsvLineReadResult.Line;

        var start = _position;
        Span<StandardCsvFieldSpan> fields = stackalloc StandardCsvFieldSpan[StandardQuotedFieldSpanCapacity];
        var firstFieldLength = 0;
        var fieldStart = start;
        foreach (var delimiterIndex in delimiterIndexesBeforeQuote)
        {
            if (delimiterIndex < fieldStart || delimiterIndex >= quoteIndex)
            {
                return false;
            }

            var field = new StandardCsvFieldSpan(
                fieldStart,
                delimiterIndex,
                delimiterIndex - fieldStart,
                hasEscapedQuotes: false,
                firstEscapedQuote: -1);
            if (!TryAddStandardField(fields, ref fieldCount, field, ref firstFieldLength))
            {
                return false;
            }

            fieldStart = delimiterIndex + 1;
        }

        if (fieldStart != quoteIndex)
        {
            return false;
        }

        var index = quoteIndex;
        if (!TryParsePrefixedStandardQuotedField(ref index, out var quotedField) ||
            !TryAddStandardField(fields, ref fieldCount, quotedField, ref firstFieldLength))
        {
            return false;
        }

        if (!TryParseStandardQuotedFieldSpanTail(
                ref index,
                delimiter,
                fields,
                ref fieldCount,
                ref firstFieldLength,
                out var recordEnd))
        {
            return false;
        }

        var emit = emitFields && (allowEmpty || recordEnd > start);
        if (emit)
        {
            VisitStandardQuotedFieldSpans(fields.Slice(0, fieldCount), recordIndex, ref fieldVisitor);
        }

        isEmptyRecord = fieldCount == 1 && firstFieldLength == 0;
        _position = recordEnd;
        ConsumeLineSeparator(_buffer[recordEnd], out separator);
        readResult = CsvLineReadResult.UnquotedRecord;
        return true;
    }

    private bool TryParseStandardQuotedFieldSpans(
        int start,
        char delimiter,
        bool trim,
        Span<StandardCsvFieldSpan> fields,
        out int fieldCount,
        out int firstFieldLength,
        out int recordEnd)
    {
        fieldCount = 0;
        firstFieldLength = 0;
        recordEnd = -1;
        var index = start;
        var pendingTrailingField = false;

        while (index < _length)
        {
            var value = _buffer[index];
            if (value == '\r' || value == '\n')
            {
                if (pendingTrailingField && !TryAddStandardField(fields, ref fieldCount, StandardCsvFieldSpan.Empty, ref firstFieldLength))
                {
                    return false;
                }

                recordEnd = index;
                return true;
            }

            if (value == delimiter)
            {
                if (!TryAddStandardField(fields, ref fieldCount, StandardCsvFieldSpan.Empty, ref firstFieldLength))
                {
                    return false;
                }

                index++;
                pendingTrailingField = true;
                continue;
            }

            StandardCsvFieldSpan field;
            if (value == '"')
            {
                if (!TryParseStandardQuotedField(ref index, delimiter, trim, out field))
                {
                    return false;
                }
            }
            else
            {
                if (!TryParseStandardUnquotedField(ref index, delimiter, trim, out field))
                {
                    return false;
                }
            }

            if (!TryAddStandardField(fields, ref fieldCount, field, ref firstFieldLength))
            {
                return false;
            }

            pendingTrailingField = false;
            if (index >= _length)
            {
                return false;
            }

            value = _buffer[index];
            if (value == delimiter)
            {
                index++;
                pendingTrailingField = true;
                continue;
            }

            if (value == '\r' || value == '\n')
            {
                recordEnd = index;
                return true;
            }

            return false;
        }

        return false;
    }

    private bool TryParseStandardQuotedField(ref int index, char delimiter, bool trim, out StandardCsvFieldSpan field)
    {
        index++;
        var valueStart = index;
        var escapeCount = 0;
        var firstEscapedQuote = -1;

        while (index < _length)
        {
            if (_buffer[index] != '"')
            {
                index++;
                continue;
            }

            if (index + 1 < _length && _buffer[index + 1] == '"')
            {
                if (firstEscapedQuote < 0)
                {
                    firstEscapedQuote = index;
                }

                escapeCount++;
                index += 2;
                continue;
            }

            var valueEnd = index;
            index++;
            if (trim)
            {
                while (index < _length &&
                       _buffer[index] != delimiter &&
                       _buffer[index] != '\r' &&
                       _buffer[index] != '\n' &&
                       char.IsWhiteSpace(_buffer[index]))
                {
                    index++;
                }
            }

            field = new StandardCsvFieldSpan(valueStart, valueEnd, valueEnd - valueStart - escapeCount, escapeCount != 0, firstEscapedQuote);
            return index < _length;
        }

        field = default;
        return false;
    }

    private bool TryParsePrefixedStandardQuotedField(ref int index, out StandardCsvFieldSpan field)
    {
        index++;
        var valueStart = index;
        var escapeCount = 0;
        var firstEscapedQuote = -1;

        while (index < _length)
        {
            var quoteOffset = _buffer.AsSpan(index, _length - index).IndexOf('"');
            if (quoteOffset < 0)
            {
                field = default;
                return false;
            }

            index += quoteOffset;
            if (index + 1 < _length && _buffer[index + 1] == '"')
            {
                if (firstEscapedQuote < 0)
                {
                    firstEscapedQuote = index;
                }

                escapeCount++;
                index += 2;
                continue;
            }

            var valueEnd = index;
            index++;
            field = new StandardCsvFieldSpan(valueStart, valueEnd, valueEnd - valueStart - escapeCount, escapeCount != 0, firstEscapedQuote);
            return index < _length;
        }

        field = default;
        return false;
    }

    private bool TryParseStandardUnquotedField(
        ref int index,
        char delimiter,
        bool trim,
        out StandardCsvFieldSpan field)
    {
        var start = index;
        while (index < _length)
        {
            var value = _buffer[index];
            if (value == '"')
            {
                field = default;
                return false;
            }

            if (value == delimiter || value == '\r' || value == '\n')
            {
                break;
            }

            index++;
        }

        var valueStart = start;
        var valueEnd = index;
        if (trim)
        {
            while (valueStart < valueEnd && char.IsWhiteSpace(_buffer[valueStart]))
            {
                valueStart++;
            }

            while (valueEnd > valueStart && char.IsWhiteSpace(_buffer[valueEnd - 1]))
            {
                valueEnd--;
            }
        }

        field = new StandardCsvFieldSpan(valueStart, valueEnd, valueEnd - valueStart, hasEscapedQuotes: false, firstEscapedQuote: -1);
        return index < _length;
    }

    private bool TryParseStandardQuotedFieldSpanTail(
        ref int index,
        char delimiter,
        Span<StandardCsvFieldSpan> fields,
        ref int fieldCount,
        ref int firstFieldLength,
        out int recordEnd)
    {
        recordEnd = -1;
        var pendingTrailingField = false;

        while (index < _length)
        {
            var value = _buffer[index];
            if (value == '\r' || value == '\n')
            {
                if (pendingTrailingField && !TryAddStandardField(fields, ref fieldCount, StandardCsvFieldSpan.Empty, ref firstFieldLength))
                {
                    return false;
                }

                recordEnd = index;
                return true;
            }

            if (value == delimiter)
            {
                index++;
                pendingTrailingField = true;
                continue;
            }

            if (!pendingTrailingField)
            {
                return false;
            }

            StandardCsvFieldSpan field;
            if (value == '"')
            {
                if (!TryParseStandardQuotedField(ref index, delimiter, trim: false, out field))
                {
                    return false;
                }
            }
            else
            {
                if (!TryParseStandardUnquotedField(ref index, delimiter, trim: false, out field))
                {
                    return false;
                }
            }

            if (!TryAddStandardField(fields, ref fieldCount, field, ref firstFieldLength))
            {
                return false;
            }

            pendingTrailingField = false;
        }

        return false;
    }

    private void VisitStandardQuotedFieldSpans<TVisitor>(
        ReadOnlySpan<StandardCsvFieldSpan> fields,
        int recordIndex,
        ref TVisitor fieldVisitor)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        for (var fieldIndex = 0; fieldIndex < fields.Length; fieldIndex++)
        {
            var field = fields[fieldIndex];
            ReadOnlySpan<char> value = _buffer.AsSpan(field.Start, field.End - field.Start);
            if (field.HasEscapedQuotes)
            {
                if (fieldVisitor.TryVisitEscapedField(recordIndex, fieldIndex, value, field.Length))
                {
                    continue;
                }

                if (field.Length >= DirectUnescapeQuotedFieldMinimumLength)
                {
                    fieldVisitor.VisitFieldValue(recordIndex, fieldIndex, CreateUnescapedQuotedField(field));
                    continue;
                }

                value = CompactEscapedQuotedField(field.Start, field.End, field.FirstEscapedQuote);
            }

            fieldVisitor.VisitField(recordIndex, fieldIndex, value);
        }
    }

    private string CreateUnescapedQuotedField(StandardCsvFieldSpan field)
    {
        return string.Create(field.Length, (this, field), static (destination, state) =>
        {
            var reader = state.Item1;
            var source = reader._buffer;
            var fieldValue = state.field;
            var readIndex = fieldValue.Start;
            var end = fieldValue.End;
            var writeIndex = 0;

            while (readIndex < end)
            {
                var quoteIndex = Array.IndexOf(source, '"', readIndex, end - readIndex);
                if (quoteIndex < 0 || quoteIndex + 1 >= end || source[quoteIndex + 1] != '"')
                {
                    var segmentLength = end - readIndex;
                    source.AsSpan(readIndex, segmentLength).CopyTo(destination[writeIndex..]);
                    break;
                }

                if (quoteIndex > readIndex)
                {
                    var segmentLength = quoteIndex - readIndex;
                    source.AsSpan(readIndex, segmentLength).CopyTo(destination[writeIndex..]);
                    writeIndex += segmentLength;
                }

                destination[writeIndex++] = '"';
                readIndex = quoteIndex + 2;
            }
        });
    }

    private ReadOnlySpan<char> CompactEscapedQuotedField(int start, int end, int firstEscapedQuote)
    {
        var index = firstEscapedQuote >= start && firstEscapedQuote < end ? firstEscapedQuote : start;
        var segmentStart = index;
        var writeIndex = index;

        while (index < end)
        {
            if (_buffer[index] != '"' || index + 1 >= end || _buffer[index + 1] != '"')
            {
                index++;
                continue;
            }

            if (index > segmentStart)
            {
                var segmentLength = index - segmentStart;
                Array.Copy(_buffer, segmentStart, _buffer, writeIndex, segmentLength);
                writeIndex += segmentLength;
            }

            _buffer[writeIndex++] = '"';
            index += 2;
            segmentStart = index;
        }

        if (index > segmentStart)
        {
            var segmentLength = index - segmentStart;
            Array.Copy(_buffer, segmentStart, _buffer, writeIndex, segmentLength);
            writeIndex += segmentLength;
        }

        return _buffer.AsSpan(start, writeIndex - start);
    }

    private static bool TryAddStandardField(
        Span<StandardCsvFieldSpan> fields,
        ref int fieldCount,
        StandardCsvFieldSpan field,
        ref int firstFieldLength)
    {
        if (fieldCount >= fields.Length)
        {
            return false;
        }

        if (fieldCount == 0)
        {
            firstFieldLength = field.Length;
        }

        fields[fieldCount++] = field;
        return true;
    }

    private readonly struct StandardCsvFieldSpan
    {
        public static readonly StandardCsvFieldSpan Empty = new(0, 0, 0, hasEscapedQuotes: false, firstEscapedQuote: -1);

        public StandardCsvFieldSpan(int start, int end, int length, bool hasEscapedQuotes, int firstEscapedQuote)
        {
            Start = start;
            End = end;
            Length = length;
            HasEscapedQuotes = hasEscapedQuotes;
            FirstEscapedQuote = firstEscapedQuote;
        }

        public int Start { get; }

        public int End { get; }

        public int Length { get; }

        public bool HasEscapedQuotes { get; }

        public int FirstEscapedQuote { get; }
    }
#endif
}
