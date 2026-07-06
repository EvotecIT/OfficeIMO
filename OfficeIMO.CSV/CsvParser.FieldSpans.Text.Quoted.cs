#nullable enable

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
#if NET8_0_OR_GREATER
    private const int TextStandardQuotedFieldSpanCapacity = 64;
    private const int TextQuotedPrefixReuseMinimumDelimiterCount = 4;

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
            var value = field.HasEscapedQuotes
                ? UnescapeTextQuotedField(text.Slice(field.Start, field.End - field.Start), field.FirstEscapedQuote - field.Start, field.Length, ref scratch)
                : text.Slice(field.Start, field.Length);
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
