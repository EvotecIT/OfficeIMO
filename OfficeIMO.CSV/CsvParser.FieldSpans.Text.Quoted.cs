#nullable enable

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
#if NET8_0_OR_GREATER
    private const int TextStandardQuotedFieldSpanCapacity = 64;
    private const int TextQuotedPrefixReuseMinimumDelimiterCount = 4;

    private static bool TryReadTextFinalQuotedRecordFieldSpansFromPrefix<TVisitor>(
        ReadOnlySpan<char> text,
        char delimiter,
        bool emitFields,
        int recordIndex,
        ReadOnlySpan<int> delimiterIndexesBeforeQuote,
        int quoteIndex,
        ref int position,
        ref TVisitor fieldVisitor,
        ref char[]? scratch,
        out int fieldCount,
        out int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        fieldCount = 0;
        firstFieldLength = 0;
        var start = position;
        if (delimiterIndexesBeforeQuote.Length == 0 ||
            delimiterIndexesBeforeQuote[^1] + 1 != quoteIndex)
        {
            return false;
        }

        var quotedPosition = quoteIndex + 1;
        var valueStart = quotedPosition;
        var quoteOffset = text.Slice(quotedPosition).IndexOf('"');
        if (quoteOffset < 0)
        {
            return false;
        }

        quotedPosition += quoteOffset;
        if (quotedPosition + 1 >= text.Length || text[quotedPosition + 1] != '"')
        {
            return CompleteTextFinalUnescapedQuotedRecordFieldSpansFromPrefix(
                text,
                emitFields,
                recordIndex,
                delimiterIndexesBeforeQuote,
                start,
                valueStart,
                quotedPosition,
                ref position,
                ref fieldVisitor,
                out fieldCount,
                out firstFieldLength);
        }

        var escapeCount = 0;
        var firstEscapedQuote = quotedPosition;
        while (quotedPosition < text.Length)
        {
            quoteOffset = text.Slice(quotedPosition).IndexOf('"');
            if (quoteOffset < 0)
            {
                return false;
            }

            quotedPosition += quoteOffset;
            if (quotedPosition + 1 < text.Length && text[quotedPosition + 1] == '"')
            {
                escapeCount++;
                quotedPosition += 2;
                continue;
            }

            break;
        }

        if (quotedPosition >= text.Length)
        {
            return false;
        }

        var valueEnd = quotedPosition;
        var nextPosition = quotedPosition + 1;
        if (nextPosition < text.Length)
        {
            var separator = text[nextPosition];
            if (separator == '\r')
            {
                nextPosition++;
                if (nextPosition < text.Length && text[nextPosition] == '\n')
                {
                    nextPosition++;
                }
            }
            else if (separator == '\n')
            {
                nextPosition++;
            }
            else
            {
                return false;
            }
        }

        var fieldStart = start;
        foreach (var delimiterIndex in delimiterIndexesBeforeQuote)
        {
            VisitTextField(text.Slice(fieldStart, delimiterIndex - fieldStart), emitFields, recordIndex, fieldCount, ref fieldVisitor, ref firstFieldLength);
            fieldCount++;
            fieldStart = delimiterIndex + 1;
        }

        var field = text.Slice(valueStart, valueEnd - valueStart);
        var fieldLength = field.Length - escapeCount;
        if (fieldCount == 0)
        {
            firstFieldLength = fieldLength;
        }

        if (emitFields)
        {
            if (escapeCount == 0)
            {
                fieldVisitor.VisitField(recordIndex, fieldCount, field);
            }
            else if (!fieldVisitor.TryVisitEscapedField(recordIndex, fieldCount, field, fieldLength))
            {
                var unescaped = UnescapeTextQuotedField(field, firstEscapedQuote - valueStart, fieldLength, ref scratch);
                fieldVisitor.VisitField(recordIndex, fieldCount, unescaped);
            }
        }

        fieldCount++;
        position = nextPosition;
        return true;
    }

    private static bool CompleteTextFinalUnescapedQuotedRecordFieldSpansFromPrefix<TVisitor>(
        ReadOnlySpan<char> text,
        bool emitFields,
        int recordIndex,
        ReadOnlySpan<int> delimiterIndexesBeforeQuote,
        int start,
        int valueStart,
        int valueEnd,
        ref int position,
        ref TVisitor fieldVisitor,
        out int fieldCount,
        out int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        var nextPosition = valueEnd + 1;
        if (nextPosition < text.Length)
        {
            var separator = text[nextPosition];
            if (separator == '\n')
            {
                nextPosition++;
            }
            else if (separator == '\r')
            {
                nextPosition++;
                if (nextPosition < text.Length && text[nextPosition] == '\n')
                {
                    nextPosition++;
                }
            }
            else
            {
                fieldCount = 0;
                firstFieldLength = 0;
                return false;
            }
        }

        fieldCount = 0;
        firstFieldLength = 0;
        var fieldStart = start;
        foreach (var delimiterIndex in delimiterIndexesBeforeQuote)
        {
            VisitTextField(text.Slice(fieldStart, delimiterIndex - fieldStart), emitFields, recordIndex, fieldCount, ref fieldVisitor, ref firstFieldLength);
            fieldCount++;
            fieldStart = delimiterIndex + 1;
        }

        var field = text.Slice(valueStart, valueEnd - valueStart);
        if (fieldCount == 0)
        {
            firstFieldLength = field.Length;
        }

        if (emitFields)
        {
            fieldVisitor.VisitField(recordIndex, fieldCount, field);
        }

        fieldCount++;
        position = nextPosition;
        return true;
    }

    private static bool TryReadTextQuotedRecordFieldSpansFromPrefix<TVisitor>(
        ReadOnlySpan<char> text,
        char delimiter,
        bool emitFields,
        int recordIndex,
        ReadOnlySpan<int> delimiterIndexesBeforeQuote,
        int quoteIndex,
        ref int position,
        ref TVisitor fieldVisitor,
        ref char[]? scratch,
        out int fieldCount,
        out int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        fieldCount = 0;
        firstFieldLength = 0;
        var start = position;
        var fieldStart = start;

        foreach (var delimiterIndex in delimiterIndexesBeforeQuote)
        {
            if (delimiterIndex < fieldStart || delimiterIndex >= quoteIndex)
            {
                return false;
            }

            fieldStart = delimiterIndex + 1;
        }

        if (fieldStart != quoteIndex)
        {
            return false;
        }

        fieldStart = start;
        foreach (var delimiterIndex in delimiterIndexesBeforeQuote)
        {
            VisitTextField(text.Slice(fieldStart, delimiterIndex - fieldStart), emitFields, recordIndex, fieldCount, ref fieldVisitor, ref firstFieldLength);
            fieldCount++;
            fieldStart = delimiterIndex + 1;
        }

        position = quoteIndex;
        if (!TryVisitTextQuotedField(text, delimiter, trim: false, emitFields, recordIndex, fieldCount, ref position, ref fieldVisitor, ref scratch, out var quotedLength))
        {
            throw new CsvParseException("Unterminated quoted field.", 0);
        }

        if (fieldCount == 0)
        {
            firstFieldLength = quotedLength;
        }

        fieldCount++;
        var pendingTrailingField = false;
        if (!TryConsumeTextQuotedFieldSeparator(text, delimiter, ref position, ref pendingTrailingField, out var recordEnded))
        {
            throw new CsvParseException("Unexpected character after CSV field.", 0);
        }

        if (recordEnded)
        {
            return true;
        }

        while (position < text.Length)
        {
            var value = text[position];
            if (value == delimiter)
            {
                if (pendingTrailingField)
                {
                    VisitTextField(text.Slice(position, 0), emitFields, recordIndex, fieldCount, ref fieldVisitor, ref firstFieldLength);
                    fieldCount++;
                }

                position++;
                pendingTrailingField = true;
                continue;
            }

            if (value == '\r' || value == '\n')
            {
                if (pendingTrailingField)
                {
                    VisitTextField(text.Slice(position, 0), emitFields, recordIndex, fieldCount, ref fieldVisitor, ref firstFieldLength);
                    fieldCount++;
                }

                ConsumeTextLineSeparator(text, ref position);
                return true;
            }

            pendingTrailingField = false;
            if (value == '"')
            {
                if (!TryVisitTextQuotedField(text, delimiter, trim: false, emitFields, recordIndex, fieldCount, ref position, ref fieldVisitor, ref scratch, out quotedLength))
                {
                    throw new CsvParseException("Unterminated quoted field.", 0);
                }

                if (fieldCount == 0)
                {
                    firstFieldLength = quotedLength;
                }

                if (!TryConsumeTextQuotedFieldSeparator(text, delimiter, ref position, ref pendingTrailingField, out recordEnded))
                {
                    throw new CsvParseException("Unexpected character after CSV field.", 0);
                }

                fieldCount++;
                if (recordEnded)
                {
                    return true;
                }

                continue;
            }
            else
            {
                var field = ReadTextUnquotedField(text, delimiter, trim: false, ref position);
                VisitTextField(field, emitFields, recordIndex, fieldCount, ref fieldVisitor, ref firstFieldLength);
            }

            fieldCount++;
            pendingTrailingField = false;
        }

        if (pendingTrailingField)
        {
            VisitTextField(text.Slice(text.Length, 0), emitFields, recordIndex, fieldCount, ref fieldVisitor, ref firstFieldLength);
            fieldCount++;
        }

        return true;
    }

    private static bool TryConsumeTextQuotedFieldSeparator(
        ReadOnlySpan<char> text,
        char delimiter,
        ref int position,
        ref bool pendingTrailingField,
        out bool recordEnded)
    {
        recordEnded = false;
        if (position >= text.Length)
        {
            recordEnded = true;
            return true;
        }

        var value = text[position];
        if (value == delimiter)
        {
            position++;
            pendingTrailingField = true;
            return true;
        }

        if (value == '\r' || value == '\n')
        {
            ConsumeTextLineSeparator(text, ref position);
            recordEnded = true;
            return true;
        }

        return false;
    }

    private static bool TryReadTextStandardQuotedRecordFieldSpansFromPrefix<TVisitor>(
        ReadOnlySpan<char> text,
        char delimiter,
        bool allowEmpty,
        bool emitFields,
        int recordIndex,
        ReadOnlySpan<int> delimiterIndexesBeforeQuote,
        int quoteIndex,
        ref int position,
        ref TVisitor fieldVisitor,
        ref char[]? scratch,
        out int fieldCount,
        out int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        fieldCount = 0;
        firstFieldLength = 0;
        var start = position;
        Span<TextStandardFieldSpan> fields = stackalloc TextStandardFieldSpan[TextStandardQuotedFieldSpanCapacity];
        var fieldStart = start;

        foreach (var delimiterIndex in delimiterIndexesBeforeQuote)
        {
            if (delimiterIndex < fieldStart || delimiterIndex >= quoteIndex)
            {
                return false;
            }

            var field = new TextStandardFieldSpan(
                fieldStart,
                delimiterIndex,
                delimiterIndex - fieldStart,
                hasEscapedQuotes: false,
                firstEscapedQuote: -1);
            if (!TryAddTextStandardField(fields, ref fieldCount, field, ref firstFieldLength))
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
        if (!TryParseTextStandardQuotedField(text, delimiter, ref index, out var quotedField) ||
            !TryAddTextStandardField(fields, ref fieldCount, quotedField, ref firstFieldLength) ||
            !TryParseTextStandardQuotedFieldSpanTail(text, delimiter, ref index, fields, ref fieldCount, ref firstFieldLength, out var recordEnd))
        {
            return false;
        }

        if (emitFields && (allowEmpty || recordEnd > start))
        {
            VisitTextStandardFieldSpans(text, fields.Slice(0, fieldCount), recordIndex, ref fieldVisitor, ref scratch);
        }

        position = recordEnd;
        if (position < text.Length)
        {
            ConsumeTextLineSeparator(text, ref position);
        }

        return true;
    }

    private static bool TryParseTextStandardQuotedFieldSpanTail(
        ReadOnlySpan<char> text,
        char delimiter,
        ref int index,
        Span<TextStandardFieldSpan> fields,
        ref int fieldCount,
        ref int firstFieldLength,
        out int recordEnd)
    {
        recordEnd = -1;
        var pendingTrailingField = false;

        while (index < text.Length)
        {
            var value = text[index];
            if (value == '\r' || value == '\n')
            {
                if (pendingTrailingField &&
                    !TryAddTextStandardField(fields, ref fieldCount, new TextStandardFieldSpan(index, index, 0, hasEscapedQuotes: false, firstEscapedQuote: -1), ref firstFieldLength))
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

            TextStandardFieldSpan field;
            if (value == '"')
            {
                if (!TryParseTextStandardQuotedField(text, delimiter, ref index, out field))
                {
                    return false;
                }
            }
            else if (!TryParseTextStandardUnquotedField(text, delimiter, ref index, out field))
            {
                return false;
            }

            if (!TryAddTextStandardField(fields, ref fieldCount, field, ref firstFieldLength))
            {
                return false;
            }

            pendingTrailingField = false;
        }

        if (pendingTrailingField &&
            !TryAddTextStandardField(fields, ref fieldCount, new TextStandardFieldSpan(text.Length, text.Length, 0, hasEscapedQuotes: false, firstEscapedQuote: -1), ref firstFieldLength))
        {
            return false;
        }

        recordEnd = text.Length;
        return fieldCount != 0;
    }

    private static bool TryParseTextStandardQuotedField(
        ReadOnlySpan<char> text,
        char delimiter,
        ref int index,
        out TextStandardFieldSpan field)
    {
        index++;
        var valueStart = index;
        var escapeCount = 0;
        var firstEscapedQuote = -1;

        while (index < text.Length)
        {
            var quoteOffset = text.Slice(index).IndexOf('"');
            if (quoteOffset < 0)
            {
                field = default;
                return false;
            }

            index += quoteOffset;
            if (index + 1 < text.Length && text[index + 1] == '"')
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
            if (index < text.Length &&
                text[index] != delimiter &&
                text[index] != '\r' &&
                text[index] != '\n')
            {
                field = default;
                return false;
            }

            field = new TextStandardFieldSpan(valueStart, valueEnd, valueEnd - valueStart - escapeCount, escapeCount != 0, firstEscapedQuote);
            return true;
        }

        field = default;
        return false;
    }

    private static bool TryParseTextStandardUnquotedField(
        ReadOnlySpan<char> text,
        char delimiter,
        ref int index,
        out TextStandardFieldSpan field)
    {
        var start = index;
        var remaining = text.Slice(index);
        var terminatorOffset = delimiter switch
        {
            ',' => remaining.IndexOfAny(CommaTextFieldTerminators),
            ';' => remaining.IndexOfAny(SemicolonTextFieldTerminators),
            '\t' => remaining.IndexOfAny(TabTextFieldTerminators),
            _ => remaining.IndexOfAny(new[] { delimiter, '\r', '\n', '"' })
        };

        if (terminatorOffset < 0)
        {
            index = text.Length;
        }
        else
        {
            index += terminatorOffset;
            if (text[index] == '"')
            {
                field = default;
                return false;
            }
        }

        field = new TextStandardFieldSpan(start, index, index - start, hasEscapedQuotes: false, firstEscapedQuote: -1);
        return true;
    }

    private static void VisitTextStandardFieldSpans<TVisitor>(
        ReadOnlySpan<char> text,
        ReadOnlySpan<TextStandardFieldSpan> fields,
        int recordIndex,
        ref TVisitor fieldVisitor,
        ref char[]? scratch)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        for (var fieldIndex = 0; fieldIndex < fields.Length; fieldIndex++)
        {
            var field = fields[fieldIndex];
            var value = text.Slice(field.Start, field.End - field.Start);
            if (field.HasEscapedQuotes)
            {
                if (fieldVisitor.TryVisitEscapedField(recordIndex, fieldIndex, value, field.Length))
                {
                    continue;
                }

                var unescaped = UnescapeTextQuotedField(value, field.FirstEscapedQuote - field.Start, field.Length, ref scratch);
                fieldVisitor.VisitField(recordIndex, fieldIndex, unescaped);
                continue;
            }

            fieldVisitor.VisitField(recordIndex, fieldIndex, value);
        }
    }

    private static bool TryAddTextStandardField(
        Span<TextStandardFieldSpan> fields,
        ref int fieldCount,
        TextStandardFieldSpan field,
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

    private readonly struct TextStandardFieldSpan
    {
        public TextStandardFieldSpan(int start, int end, int length, bool hasEscapedQuotes, int firstEscapedQuote)
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
