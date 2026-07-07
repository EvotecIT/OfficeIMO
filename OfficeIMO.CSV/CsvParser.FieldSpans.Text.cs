#nullable enable

using System.Buffers;

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
#if NET8_0_OR_GREATER
    private static readonly SearchValues<char> CommaTextFieldTerminators = SearchValues.Create(",\r\n\"");
    private static readonly SearchValues<char> SemicolonTextFieldTerminators = SearchValues.Create(";\r\n\"");
    private static readonly SearchValues<char> TabTextFieldTerminators = SearchValues.Create("\t\r\n\"");

    internal static void ReadFieldSpans<TVisitor>(
        ReadOnlySpan<char> text,
        CsvLoadOptions options,
        int recordsToSkip,
        ref TVisitor fieldVisitor)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        var delimiter = options.Delimiter;
        var trim = options.TrimWhitespace;
        var allowEmpty = options.AllowEmptyLines;
        var position = 0;
        var recordIndex = 0;
        var emittedRecordCount = 0;
        var useAvx2UnquotedFastPath = true;
        char[]? scratch = null;

        try
        {
            while (position < text.Length)
            {
                var recordStart = position;
                if (TrySkipTextEmptyRecord(text, trim, allowEmpty, ref position))
                {
                    continue;
                }

                var skipCommentRecord = options.SkipCommentRows &&
                    text[position] == options.CommentCharacter &&
                    !CanReadW3CFieldsHeader(options, emittedRecordCount);
                if (recordsToSkip > 0 &&
                    !skipCommentRecord &&
                    !trim &&
                    TrySkipTextUnquotedRecord(text, ref position))
                {
                    recordsToSkip--;
                    continue;
                }

                var emitFields = recordsToSkip == 0 && !skipCommentRecord;
                int fieldCount;
                int firstFieldLength;
                if (skipCommentRecord ||
                    !TryReadTextUnquotedRecordFieldSpans(
                        text,
                        delimiter,
                        trim,
                        allowEmpty,
                        emitFields,
                        recordIndex,
                        ref useAvx2UnquotedFastPath,
                        ref position,
                        ref fieldVisitor,
                        ref scratch,
                        out fieldCount,
                        out firstFieldLength))
                {
                    fieldCount = ReadTextRecordFieldSpans(
                        text,
                        delimiter,
                        trim,
                        emitFields,
                        recordIndex,
                        ref position,
                        ref fieldVisitor,
                        ref scratch,
                        out firstFieldLength);
                }

                var isEmptyRecord = fieldCount == 1 && firstFieldLength == 0;
                var shouldEmit = fieldCount != 0 && (allowEmpty || !isEmptyRecord);
                if (!shouldEmit || skipCommentRecord)
                {
                    continue;
                }

                if (recordsToSkip > 0)
                {
                    recordsToSkip--;
                    continue;
                }

                recordIndex++;
                emittedRecordCount++;

                if (position == recordStart)
                {
                    break;
                }
            }
        }
        finally
        {
            if (scratch != null)
            {
                ArrayPool<char>.Shared.Return(scratch);
            }
        }
    }

    private static bool TryReadTextUnquotedRecordFieldSpans<TVisitor>(
        ReadOnlySpan<char> text,
        char delimiter,
        bool trim,
        bool allowEmpty,
        bool emitFields,
        int recordIndex,
        ref bool useAvx2UnquotedFastPath,
        ref int position,
        ref TVisitor fieldVisitor,
        ref char[]? scratch,
        out int fieldCount,
        out int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        fieldCount = 0;
        firstFieldLength = 0;

#if NET8_0_OR_GREATER
        var encounteredQuote = false;
        if (useAvx2UnquotedFastPath &&
            !trim &&
            delimiter <= byte.MaxValue &&
            System.Runtime.Intrinsics.X86.Avx2.IsSupported &&
            TryReadTextUnquotedRecordFieldSpansAvx2(
                text,
                delimiter,
                allowEmpty,
                emitFields,
                recordIndex,
                ref position,
                ref fieldVisitor,
                ref scratch,
                out fieldCount,
                out firstFieldLength,
                out encounteredQuote))
        {
            return true;
        }

        if (encounteredQuote)
        {
            useAvx2UnquotedFastPath = false;
        }
#endif

        var start = position;
        var specialOffset = text.Slice(start).IndexOfAny('"', '\r', '\n');
        var endsAtTextEnd = false;
        int recordEnd;
        if (specialOffset < 0)
        {
            recordEnd = text.Length;
            endsAtTextEnd = true;
        }
        else
        {
            recordEnd = start + specialOffset;
            if (text[recordEnd] == '"')
            {
                return false;
            }
        }

        var emitNonEmptyRecord = allowEmpty || (trim
            ? HasTextNonWhitespaceOrDelimiter(text.Slice(start, recordEnd - start), delimiter)
            : recordEnd != start);
        fieldCount = VisitTextUnquotedFieldSpans(
            text,
            start,
            recordEnd,
            delimiter,
            trim,
            emitNonEmptyRecord && emitFields,
            recordIndex,
            ref fieldVisitor,
            out firstFieldLength);
        position = recordEnd;
        if (!endsAtTextEnd)
        {
            ConsumeTextLineSeparator(text, ref position);
        }

        return true;
    }

    private static bool TryReadTextUnquotedRecordFieldSpansAvx2<TVisitor>(
        ReadOnlySpan<char> text,
        char delimiter,
        bool allowEmpty,
        bool emitFields,
        int recordIndex,
        ref int position,
        ref TVisitor fieldVisitor,
        ref char[]? scratch,
        out int fieldCount,
        out int firstFieldLength,
        out bool encounteredQuote)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        fieldCount = 0;
        firstFieldLength = 0;
        encounteredQuote = false;

        var start = position;
        var end = text.Length - 32;
        if (start > end)
        {
            return false;
        }

        Span<int> delimiterIndexes = stackalloc int[64];
        var delimiterCount = 0;
        var pos = start;
        var delimiterVector = System.Runtime.Intrinsics.Vector256.Create((byte)delimiter);
        var quoteVector = System.Runtime.Intrinsics.Vector256.Create((byte)'"');
        var carriageReturnVector = System.Runtime.Intrinsics.Vector256.Create((byte)'\r');
        var lineFeedVector = System.Runtime.Intrinsics.Vector256.Create((byte)'\n');

        while (pos <= end)
        {
            var values = System.Runtime.InteropServices.MemoryMarshal.Cast<char, short>(text.Slice(pos, 32));
            var first = System.Runtime.Intrinsics.Vector256.LoadUnsafe(ref System.Runtime.InteropServices.MemoryMarshal.GetReference(values));
            var second = System.Runtime.Intrinsics.Vector256.LoadUnsafe(ref System.Runtime.InteropServices.MemoryMarshal.GetReference(values.Slice(16)));
            var packed = System.Runtime.Intrinsics.X86.Avx2.PackUnsignedSaturate(first, second);
            var packedBytes = System.Runtime.Intrinsics.Vector256.AsByte(
                System.Runtime.Intrinsics.X86.Avx2.Permute4x64(System.Runtime.Intrinsics.Vector256.AsInt64(packed), 0b11_01_10_00));

            var delimiterMask = (uint)System.Runtime.Intrinsics.X86.Avx2.MoveMask(
                System.Runtime.Intrinsics.X86.Avx2.CompareEqual(packedBytes, delimiterVector));
            var quoteMask = (uint)System.Runtime.Intrinsics.X86.Avx2.MoveMask(
                System.Runtime.Intrinsics.X86.Avx2.CompareEqual(packedBytes, quoteVector));
            var carriageReturnMask = (uint)System.Runtime.Intrinsics.X86.Avx2.MoveMask(
                System.Runtime.Intrinsics.X86.Avx2.CompareEqual(packedBytes, carriageReturnVector));
            var lineFeedMask = (uint)System.Runtime.Intrinsics.X86.Avx2.MoveMask(
                System.Runtime.Intrinsics.X86.Avx2.CompareEqual(packedBytes, lineFeedVector));
            var terminalMask = quoteMask | carriageReturnMask | lineFeedMask;

            if (terminalMask != 0)
            {
                var terminalOffset = System.Numerics.BitOperations.TrailingZeroCount(terminalMask);
                var delimiterMaskBeforeTerminal = delimiterMask & ((1u << terminalOffset) - 1u);
                if (!AddTextDelimiterIndexes(delimiterMaskBeforeTerminal, pos, delimiterIndexes, ref delimiterCount))
                {
                    return false;
                }

                if (((quoteMask >> terminalOffset) & 1u) != 0)
                {
                    encounteredQuote = true;
                    if (delimiterCount <= 2 &&
                        TryReadTextQuoteAwareRecordFieldSpansAvx2FromCurrentChunk(
                            text,
                            delimiter,
                            allowEmpty,
                            emitFields,
                            recordIndex,
                            start,
                            pos,
                            delimiterMask,
                            quoteMask,
                            carriageReturnMask,
                            lineFeedMask,
                            ref position,
                            ref fieldVisitor,
                            ref scratch,
                            out fieldCount,
                            out firstFieldLength))
                    {
                        return true;
                    }

                    var quoteIndex = pos + terminalOffset;
                    if (delimiterCount >= TextQuotedPrefixReuseMinimumDelimiterCount &&
                        (TryReadTextFinalQuotedRecordFieldSpansFromPrefix(
                            text,
                            delimiter,
                            emitFields,
                            recordIndex,
                            delimiterIndexes.Slice(0, delimiterCount),
                            quoteIndex,
                            ref position,
                            ref fieldVisitor,
                            ref scratch,
                            out fieldCount,
                            out firstFieldLength) ||
                         TryReadTextQuotedRecordFieldSpansFromPrefix(
                            text,
                            delimiter,
                            emitFields,
                            recordIndex,
                            delimiterIndexes.Slice(0, delimiterCount),
                            quoteIndex,
                            ref position,
                            ref fieldVisitor,
                            ref scratch,
                            out fieldCount,
                            out firstFieldLength) ||
                         TryReadTextStandardQuotedRecordFieldSpansFromPrefix(
                            text,
                            delimiter,
                            allowEmpty,
                            emitFields,
                            recordIndex,
                            delimiterIndexes.Slice(0, delimiterCount),
                            quoteIndex,
                            ref position,
                            ref fieldVisitor,
                            ref scratch,
                            out fieldCount,
                            out firstFieldLength)))
                    {
                        return true;
                    }

                    return false;
                }

                var recordEnd = pos + terminalOffset;
                var lineLength = recordEnd - start;
                fieldCount = VisitIndexedTextUntrimmedUnquotedFieldSpans(
                    text,
                    start,
                    recordEnd,
                    delimiterIndexes.Slice(0, delimiterCount),
                    (allowEmpty || lineLength != 0) && emitFields,
                    recordIndex,
                    ref fieldVisitor,
                    out firstFieldLength);
                position = recordEnd;
                ConsumeTextLineSeparator(text, ref position);
                return true;
            }

            if (!AddTextDelimiterIndexes(delimiterMask, pos, delimiterIndexes, ref delimiterCount))
            {
                return false;
            }

            pos += 32;
        }

        return false;
    }

    private static int VisitIndexedTextUntrimmedUnquotedFieldSpans<TVisitor>(
        ReadOnlySpan<char> text,
        int start,
        int end,
        ReadOnlySpan<int> delimiterIndexes,
        bool emit,
        int recordIndex,
        ref TVisitor fieldVisitor,
        out int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        var fieldIndex = 0;
        var fieldStart = start;
        firstFieldLength = 0;
        foreach (var delimiterIndex in delimiterIndexes)
        {
            var length = delimiterIndex - fieldStart;
            if (fieldIndex == 0)
            {
                firstFieldLength = length;
            }

            if (emit)
            {
                fieldVisitor.VisitField(recordIndex, fieldIndex, text.Slice(fieldStart, length));
            }

            fieldIndex++;
            fieldStart = delimiterIndex + 1;
        }

        var finalLength = end - fieldStart;
        if (fieldIndex == 0)
        {
            firstFieldLength = finalLength;
        }

        if (emit)
        {
            fieldVisitor.VisitField(recordIndex, fieldIndex, text.Slice(fieldStart, finalLength));
        }

        return fieldIndex + 1;
    }

    private static bool AddTextDelimiterIndexes(uint delimiterMask, int chunkStart, Span<int> delimiterIndexes, ref int delimiterCount)
    {
        while (delimiterMask != 0)
        {
            if (delimiterCount == delimiterIndexes.Length)
            {
                return false;
            }

            var offset = System.Numerics.BitOperations.TrailingZeroCount(delimiterMask);
            delimiterIndexes[delimiterCount++] = chunkStart + offset;
            delimiterMask &= delimiterMask - 1;
        }

        return true;
    }

    private static bool HasTextNonWhitespaceOrDelimiter(ReadOnlySpan<char> text, char delimiter)
    {
        foreach (var value in text)
        {
            if (value == delimiter || !char.IsWhiteSpace(value))
            {
                return true;
            }
        }

        return false;
    }

    private static int VisitTextUnquotedFieldSpans<TVisitor>(
        ReadOnlySpan<char> text,
        int start,
        int end,
        char delimiter,
        bool trim,
        bool emit,
        int recordIndex,
        ref TVisitor fieldVisitor,
        out int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        var fieldIndex = 0;
        var fieldStart = start;
        firstFieldLength = 0;
        for (var i = start; i < end; i++)
        {
            if (text[i] != delimiter)
            {
                continue;
            }

            VisitTextField(text.Slice(fieldStart, i - fieldStart), trim, emit, recordIndex, fieldIndex, ref fieldVisitor, ref firstFieldLength);
            fieldIndex++;
            fieldStart = i + 1;
        }

        VisitTextField(text.Slice(fieldStart, end - fieldStart), trim, emit, recordIndex, fieldIndex, ref fieldVisitor, ref firstFieldLength);
        return fieldIndex + 1;
    }

    private static bool TrySkipTextEmptyRecord(ReadOnlySpan<char> text, bool trim, bool allowEmpty, ref int position)
    {
        if (allowEmpty)
        {
            return false;
        }

        if (text[position] == '\r' || text[position] == '\n')
        {
            ConsumeTextLineSeparator(text, ref position);
            return true;
        }

        if (!trim)
        {
            return false;
        }

        var scan = position;
        while (scan < text.Length)
        {
            var value = text[scan];
            if (value == '\r' || value == '\n')
            {
                position = scan;
                ConsumeTextLineSeparator(text, ref position);
                return true;
            }

            if (!char.IsWhiteSpace(value))
            {
                return false;
            }

            scan++;
        }

        position = scan;
        return true;
    }

    private static bool TrySkipTextUnquotedRecord(ReadOnlySpan<char> text, ref int position)
    {
        var start = position;
        var specialOffset = text.Slice(start).IndexOfAny('"', '\r', '\n');
        if (specialOffset < 0)
        {
            position = text.Length;
            return text.Length != start;
        }

        var recordEnd = start + specialOffset;
        if (text[recordEnd] == '"' || recordEnd == start)
        {
            return false;
        }

        position = recordEnd;
        ConsumeTextLineSeparator(text, ref position);
        return true;
    }

    private static int ReadTextRecordFieldSpans<TVisitor>(
        ReadOnlySpan<char> text,
        char delimiter,
        bool trim,
        bool emitFields,
        int recordIndex,
        ref int position,
        ref TVisitor fieldVisitor,
        ref char[]? scratch,
        out int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        var fieldIndex = 0;
        var pendingTrailingField = false;
        firstFieldLength = 0;

        while (position < text.Length)
        {
            var value = text[position];
            if (value == '\r' || value == '\n')
            {
                if (pendingTrailingField || fieldIndex == 0)
                {
                    VisitTextField(text.Slice(position, 0), emitFields, recordIndex, fieldIndex, ref fieldVisitor, ref firstFieldLength);
                    fieldIndex++;
                }

                ConsumeTextLineSeparator(text, ref position);
                return fieldIndex;
            }

            if (value == delimiter)
            {
                VisitTextField(text.Slice(position, 0), emitFields, recordIndex, fieldIndex, ref fieldVisitor, ref firstFieldLength);
                fieldIndex++;
                position++;
                pendingTrailingField = true;
                continue;
            }

            pendingTrailingField = false;
            if (value == '"')
            {
                if (!TryVisitTextQuotedField(text, delimiter, trim, emitFields, recordIndex, fieldIndex, ref position, ref fieldVisitor, ref scratch, out var quotedLength))
                {
                    throw new CsvParseException("Unterminated quoted field.", 0);
                }

                if (fieldIndex == 0)
                {
                    firstFieldLength = quotedLength;
                }
            }
            else
            {
                var field = ReadTextUnquotedField(text, delimiter, trim, ref position);
                VisitTextField(field, emitFields, recordIndex, fieldIndex, ref fieldVisitor, ref firstFieldLength);
            }

            fieldIndex++;
            if (position >= text.Length)
            {
                return fieldIndex;
            }

            value = text[position];
            if (value == delimiter)
            {
                position++;
                pendingTrailingField = true;
                continue;
            }

            if (value == '\r' || value == '\n')
            {
                ConsumeTextLineSeparator(text, ref position);
                return fieldIndex;
            }

            throw new CsvParseException("Unexpected character after CSV field.", 0);
        }

        if (pendingTrailingField)
        {
            VisitTextField(text.Slice(text.Length, 0), emitFields, recordIndex, fieldIndex, ref fieldVisitor, ref firstFieldLength);
            fieldIndex++;
        }

        return fieldIndex;
    }

    private static ReadOnlySpan<char> ReadTextUnquotedField(ReadOnlySpan<char> text, char delimiter, bool trim, ref int position)
    {
        var start = position;
        var remaining = text.Slice(position);
        var terminatorOffset = delimiter switch
        {
            ',' => remaining.IndexOfAny(CommaTextFieldTerminators),
            ';' => remaining.IndexOfAny(SemicolonTextFieldTerminators),
            '\t' => remaining.IndexOfAny(TabTextFieldTerminators),
            _ => remaining.IndexOfAny(new[] { delimiter, '\r', '\n', '"' })
        };
        if (terminatorOffset < 0)
        {
            position = text.Length;
        }
        else
        {
            position += terminatorOffset;
            if (text[position] == '"')
            {
                throw new CsvParseException("Unexpected quote in unquoted CSV field.", 0);
            }
        }

        var field = text.Slice(start, position - start);
        return trim ? TrimTextField(field) : field;
    }

    private static bool TryVisitTextQuotedField<TVisitor>(
        ReadOnlySpan<char> text,
        char delimiter,
        bool trim,
        bool emitFields,
        int recordIndex,
        int fieldIndex,
        ref int position,
        ref TVisitor fieldVisitor,
        ref char[]? scratch,
        out int fieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        position++;
        var valueStart = position;
        var escapeCount = 0;
        var firstEscapedQuote = -1;

        while (position < text.Length)
        {
            var quoteOffset = text.Slice(position).IndexOf('"');
            if (quoteOffset < 0)
            {
                fieldLength = 0;
                return false;
            }

            position += quoteOffset;
            if (position + 1 < text.Length && text[position + 1] == '"')
            {
                if (firstEscapedQuote < 0)
                {
                    firstEscapedQuote = position;
                }

                escapeCount++;
                position += 2;
                continue;
            }

            var valueEnd = position;
            position++;
            if (trim)
            {
                while (position < text.Length &&
                    text[position] != delimiter &&
                    text[position] != '\r' &&
                    text[position] != '\n' &&
                    char.IsWhiteSpace(text[position]))
                {
                    position++;
                }
            }

            fieldLength = valueEnd - valueStart - escapeCount;
            if (!emitFields)
            {
                return true;
            }

            var field = text.Slice(valueStart, valueEnd - valueStart);
            if (escapeCount == 0)
            {
                fieldVisitor.VisitField(recordIndex, fieldIndex, field);
                return true;
            }

            if (fieldVisitor.TryVisitEscapedField(recordIndex, fieldIndex, field, fieldLength))
            {
                return true;
            }

            var unescaped = UnescapeTextQuotedField(field, firstEscapedQuote - valueStart, fieldLength, ref scratch);
            fieldVisitor.VisitField(recordIndex, fieldIndex, unescaped);
            return true;
        }

        fieldLength = 0;
        return false;
    }

    private static ReadOnlySpan<char> UnescapeTextQuotedField(
        ReadOnlySpan<char> field,
        int firstEscapedQuote,
        int fieldLength,
        ref char[]? scratch)
    {
        if (scratch == null || scratch.Length < fieldLength)
        {
            if (scratch != null)
            {
                ArrayPool<char>.Shared.Return(scratch);
            }

            scratch = ArrayPool<char>.Shared.Rent(fieldLength);
        }

        var readIndex = firstEscapedQuote >= 0 ? firstEscapedQuote : 0;
        if (readIndex > 0)
        {
            field.Slice(0, readIndex).CopyTo(scratch.AsSpan());
        }

        var writeIndex = readIndex;
        while (readIndex < field.Length)
        {
            var quoteOffset = field.Slice(readIndex).IndexOf('"');
            if (quoteOffset < 0)
            {
                field.Slice(readIndex).CopyTo(scratch.AsSpan(writeIndex));
                writeIndex += field.Length - readIndex;
                break;
            }

            if (quoteOffset > 0)
            {
                field.Slice(readIndex, quoteOffset).CopyTo(scratch.AsSpan(writeIndex));
                writeIndex += quoteOffset;
                readIndex += quoteOffset;
            }

            if (readIndex + 1 < field.Length && field[readIndex + 1] == '"')
            {
                scratch[writeIndex++] = '"';
                readIndex += 2;
                continue;
            }

            scratch[writeIndex++] = field[readIndex++];
        }

        return scratch.AsSpan(0, writeIndex);
    }

    private static void VisitTextField<TVisitor>(
        ReadOnlySpan<char> value,
        bool emitFields,
        int recordIndex,
        int fieldIndex,
        ref TVisitor fieldVisitor,
        ref int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        if (fieldIndex == 0)
        {
            firstFieldLength = value.Length;
        }

        if (emitFields)
        {
            fieldVisitor.VisitField(recordIndex, fieldIndex, value);
        }
    }

    private static void VisitTextField<TVisitor>(
        ReadOnlySpan<char> value,
        bool trim,
        bool emitFields,
        int recordIndex,
        int fieldIndex,
        ref TVisitor fieldVisitor,
        ref int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        if (trim)
        {
            value = TrimTextField(value);
        }

        VisitTextField(value, emitFields, recordIndex, fieldIndex, ref fieldVisitor, ref firstFieldLength);
    }

    private static ReadOnlySpan<char> TrimTextField(ReadOnlySpan<char> field)
    {
        var start = 0;
        var end = field.Length;
        while (start < end && char.IsWhiteSpace(field[start]))
        {
            start++;
        }

        while (end > start && char.IsWhiteSpace(field[end - 1]))
        {
            end--;
        }

        return field.Slice(start, end - start);
    }

    private static void ConsumeTextLineSeparator(ReadOnlySpan<char> text, ref int position)
    {
        if (text[position] == '\r' && position + 1 < text.Length && text[position + 1] == '\n')
        {
            position += 2;
            return;
        }

        position++;
    }
#endif
}
