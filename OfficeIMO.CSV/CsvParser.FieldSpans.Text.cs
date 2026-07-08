#nullable enable

using System.Buffers;

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
#if NET8_0_OR_GREATER
    private static readonly SearchValues<char> CommaTextFieldTerminators = SearchValues.Create(",\r\n\"");
    private static readonly SearchValues<char> SemicolonTextFieldTerminators = SearchValues.Create(";\r\n\"");
    private static readonly SearchValues<char> TabTextFieldTerminators = SearchValues.Create("\t\r\n\"");
    private static readonly System.Runtime.Intrinsics.Vector256<byte> QuoteByteVector = System.Runtime.Intrinsics.Vector256.Create((byte)'"');
    private static readonly System.Runtime.Intrinsics.Vector256<byte> CarriageReturnByteVector = System.Runtime.Intrinsics.Vector256.Create((byte)'\r');
    private static readonly System.Runtime.Intrinsics.Vector256<byte> LineFeedByteVector = System.Runtime.Intrinsics.Vector256.Create((byte)'\n');
    private const int TextQuoteFreeProbeMinimumLength = 64 * 1024;

    internal static void ReadFieldSpans<TVisitor>(
        ReadOnlySpan<char> text,
        CsvLoadOptions options,
        int recordsToSkip,
        ref TVisitor fieldVisitor)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        if (HasFieldLengthLimits(options))
        {
            using var reader = new StringReader(text.ToString());
            ReadFieldSpansMaterialized(reader, options, recordsToSkip, ref fieldVisitor);
            return;
        }

        if (UsesTextDelimiter(options))
        {
            using var reader = new StringReader(text.ToString());
            ReadFieldSpansTextDelimiter(reader, options, recordsToSkip, ref fieldVisitor);
            return;
        }

        if (options.ParseErrorAction == CsvParseErrorAction.SkipRow)
        {
            using var reader = new StringReader(text.ToString());
            ReadFieldSpansLineOrQuoted(reader, options, recordsToSkip, ref fieldVisitor);
            return;
        }

        var delimiter = GetDelimiterChar(options);
        var trim = options.TrimWhitespace;
        var strictQuotes = options.QuoteParsingMode == CsvQuoteParsingMode.Strict;
        var allowEmpty = options.AllowEmptyLines;
        var position = 0;
        var recordIndex = 0;
        var emittedRecordCount = 0;
        var lineNumber = 1;
        var useAvx2UnquotedFastPath = true;
        var textMayContainQuote = text.Length < TextQuoteFreeProbeMinimumLength || text.IndexOf('"') >= 0;
        var unquotedDelimiterIndexCapacity = 64;
        var projectedFieldVisitor = fieldVisitor as ICsvProjectedFieldSpanVisitor;
        char[]? scratch = null;
        var delimiterVector = System.Runtime.Intrinsics.Vector256<byte>.Zero;
        if (!trim &&
            delimiter <= byte.MaxValue &&
            System.Runtime.Intrinsics.X86.Avx2.IsSupported)
        {
            delimiterVector = System.Runtime.Intrinsics.Vector256.Create((byte)delimiter);
        }

        try
        {
            while (position < text.Length)
            {
                ThrowIfCancellationRequested(options);
                var recordStart = position;
                if (TrySkipTextEmptyRecord(text, trim, allowEmpty, ref position))
                {
                    continue;
                }

                var startsWithCommentCharacter = text[position] == options.CommentCharacter;
                var isW3CFieldsHeader = startsWithCommentCharacter &&
                    CanReadW3CFieldsHeader(options, emittedRecordCount) &&
                    IsTextW3CFieldsLine(text, position);
                var skipCommentRecord = startsWithCommentCharacter &&
                    !isW3CFieldsHeader &&
                    (options.SkipCommentRows ||
                        (options.HasHeaderRow &&
                            options.Header is null &&
                            options.SkipCommentRowsBeforeHeader &&
                            emittedRecordCount <= GetParserInitialRecordsToSkip(options)));
                if (skipCommentRecord)
                {
                    SkipTextRecord(text, ref position);
                    continue;
                }

                if (recordsToSkip > 0 &&
                    !trim &&
                    TrySkipTextUnquotedRecord(text, delimiter, ref position, out var skippedDelimiterCount))
                {
                    unquotedDelimiterIndexCapacity = GetTextDelimiterIndexCapacity(skippedDelimiterCount);
                    recordsToSkip--;
                    continue;
                }

                var emitFields = recordsToSkip == 0;
                int fieldCount;
                int firstFieldLength;
                try
                {
                    if (!TryReadTextUnquotedRecordFieldSpans(
                            text,
                            delimiter,
                            trim,
                            allowEmpty,
                            emitFields,
                            recordIndex,
                            ref useAvx2UnquotedFastPath,
                            ref unquotedDelimiterIndexCapacity,
                            textMayContainQuote,
                            delimiterVector,
                            ref position,
                            projectedFieldVisitor,
                            ref fieldVisitor,
                            ref scratch,
                            out fieldCount,
                            out firstFieldLength))
                    {
                        fieldCount = ReadTextRecordFieldSpans(
                            text,
                            delimiter,
                            trim,
                            strictQuotes,
                            emitFields,
                            recordIndex,
                            ref position,
                            projectedFieldVisitor,
                            ref fieldVisitor,
                            ref scratch,
                            out firstFieldLength);
                    }
                }
                catch (CsvParseException ex) when (HandleParseError(options, ex, lineNumber))
                {
                    position = recordStart;
                    SkipTextRecord(text, ref position);
                    lineNumber++;
                    continue;
                }

                var isEmptyRecord = fieldCount == 1 && firstFieldLength == 0;
                var shouldEmit = fieldCount != 0 && (allowEmpty || !isEmptyRecord);
                if (!shouldEmit)
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
                lineNumber++;

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
        ref int unquotedDelimiterIndexCapacity,
        bool textMayContainQuote,
        System.Runtime.Intrinsics.Vector256<byte> delimiterVector,
        ref int position,
        ICsvProjectedFieldSpanVisitor? projectedFieldVisitor,
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
            !textMayContainQuote &&
            TryReadTextQuoteFreeRecordFieldSpansAvx2(
                text,
                delimiter,
                allowEmpty,
                emitFields,
                recordIndex,
                delimiterVector,
                ref position,
                projectedFieldVisitor,
                ref fieldVisitor,
                out fieldCount,
                out firstFieldLength))
        {
            return true;
        }

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
                ref unquotedDelimiterIndexCapacity,
                delimiterVector,
                ref position,
                projectedFieldVisitor,
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
        var specialOffset = textMayContainQuote
            ? text.Slice(start).IndexOfAny('"', '\r', '\n')
            : text.Slice(start).IndexOfAny('\r', '\n');
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
            projectedFieldVisitor,
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
        ref int delimiterIndexCapacity,
        System.Runtime.Intrinsics.Vector256<byte> delimiterVector,
        ref int position,
        ICsvProjectedFieldSpanVisitor? projectedFieldVisitor,
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

        Span<int> delimiterIndexes = delimiterIndexCapacity switch
        {
            16 => stackalloc int[16],
            32 => stackalloc int[32],
            _ => stackalloc int[64],
        };
        var delimiterCount = 0;
        var pos = start;

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
                System.Runtime.Intrinsics.X86.Avx2.CompareEqual(packedBytes, QuoteByteVector));
            var carriageReturnMask = (uint)System.Runtime.Intrinsics.X86.Avx2.MoveMask(
                System.Runtime.Intrinsics.X86.Avx2.CompareEqual(packedBytes, CarriageReturnByteVector));
            var lineFeedMask = (uint)System.Runtime.Intrinsics.X86.Avx2.MoveMask(
                System.Runtime.Intrinsics.X86.Avx2.CompareEqual(packedBytes, LineFeedByteVector));
            var terminalMask = quoteMask | carriageReturnMask | lineFeedMask;

            if (terminalMask != 0)
            {
                var terminalOffset = System.Numerics.BitOperations.TrailingZeroCount(terminalMask);
                var delimiterMaskBeforeTerminal = delimiterMask & ((1u << terminalOffset) - 1u);
                if (!AddTextDelimiterIndexes(delimiterMaskBeforeTerminal, pos, delimiterIndexes, ref delimiterCount))
                {
                    delimiterIndexCapacity = 64;
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
                            delimiterVector,
                            ref position,
                            projectedFieldVisitor,
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
                            projectedFieldVisitor,
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
                            projectedFieldVisitor,
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
                            projectedFieldVisitor,
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
                    projectedFieldVisitor,
                    ref fieldVisitor,
                    out firstFieldLength);
                position = recordEnd;
                ConsumeTextLineSeparator(text, ref position);
                return true;
            }

            if (!AddTextDelimiterIndexes(delimiterMask, pos, delimiterIndexes, ref delimiterCount))
            {
                delimiterIndexCapacity = 64;
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
        ICsvProjectedFieldSpanVisitor? projectedFieldVisitor,
        ref TVisitor fieldVisitor,
        out int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        if (!emit)
        {
            firstFieldLength = delimiterIndexes.Length == 0
                ? end - start
                : delimiterIndexes[0] - start;
            return delimiterIndexes.Length + 1;
        }

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

            if (CsvFieldSpanProjection.ShouldVisitField(projectedFieldVisitor, recordIndex, fieldIndex))
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

        if (CsvFieldSpanProjection.ShouldVisitField(projectedFieldVisitor, recordIndex, fieldIndex))
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
        ICsvProjectedFieldSpanVisitor? projectedFieldVisitor,
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

            VisitTextField(text.Slice(fieldStart, i - fieldStart), trim, emit, recordIndex, fieldIndex, projectedFieldVisitor, ref fieldVisitor, ref firstFieldLength);
            fieldIndex++;
            fieldStart = i + 1;
        }

        VisitTextField(text.Slice(fieldStart, end - fieldStart), trim, emit, recordIndex, fieldIndex, projectedFieldVisitor, ref fieldVisitor, ref firstFieldLength);
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

    private static bool TrySkipTextUnquotedRecord(ReadOnlySpan<char> text, char delimiter, ref int position, out int delimiterCount)
    {
        delimiterCount = 0;
        var start = position;
        var specialOffset = text.Slice(start).IndexOfAny('"', '\r', '\n');
        if (specialOffset < 0)
        {
            delimiterCount = CountTextDelimiters(text.Slice(start), delimiter);
            position = text.Length;
            return text.Length != start;
        }

        var recordEnd = start + specialOffset;
        if (text[recordEnd] == '"' || recordEnd == start)
        {
            return false;
        }

        delimiterCount = CountTextDelimiters(text.Slice(start, recordEnd - start), delimiter);
        position = recordEnd;
        ConsumeTextLineSeparator(text, ref position);
        return true;
    }

    private static int GetTextDelimiterIndexCapacity(int delimiterCount)
    {
        if (delimiterCount <= 16)
        {
            return 16;
        }

        return delimiterCount <= 32 ? 32 : 64;
    }

    private static int CountTextDelimiters(ReadOnlySpan<char> text, char delimiter)
    {
        var delimiterCount = 0;
        foreach (var value in text)
        {
            if (value == delimiter)
            {
                delimiterCount++;
            }
        }

        return delimiterCount;
    }

    private static int ReadTextRecordFieldSpans<TVisitor>(
        ReadOnlySpan<char> text,
        char delimiter,
        bool trim,
        bool strictQuotes,
        bool emitFields,
        int recordIndex,
        ref int position,
        ICsvProjectedFieldSpanVisitor? projectedFieldVisitor,
        ref TVisitor fieldVisitor,
        ref char[]? scratch,
        out int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        var recordStart = position;
        if (!strictQuotes && ShouldUseFlexibleTextRecordParsing(text, delimiter, trim, recordStart))
        {
            return ReadFlexibleTextRecordFieldSpans(
                text,
                delimiter,
                trim,
                emitFields,
                recordIndex,
                ref position,
                projectedFieldVisitor,
                ref fieldVisitor,
                out firstFieldLength);
        }

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
                    VisitTextField(text.Slice(position, 0), emitFields, recordIndex, fieldIndex, projectedFieldVisitor, ref fieldVisitor, ref firstFieldLength);
                    fieldIndex++;
                }

                ConsumeTextLineSeparator(text, ref position);
                return fieldIndex;
            }

            if (value == delimiter)
            {
                VisitTextField(text.Slice(position, 0), emitFields, recordIndex, fieldIndex, projectedFieldVisitor, ref fieldVisitor, ref firstFieldLength);
                fieldIndex++;
                position++;
                pendingTrailingField = true;
                continue;
            }

            pendingTrailingField = false;
            if (value == '"')
            {
                if (!TryVisitTextQuotedField(text, delimiter, trim, emitFields, recordIndex, fieldIndex, ref position, projectedFieldVisitor, ref fieldVisitor, ref scratch, out var quotedLength))
                {
                    if (!strictQuotes)
                    {
                        position = recordStart;
                        return ReadFlexibleTextRecordFieldSpans(
                            text,
                            delimiter,
                            trim,
                            emitFields,
                            recordIndex,
                            ref position,
                            projectedFieldVisitor,
                            ref fieldVisitor,
                            out firstFieldLength);
                    }

                    throw new CsvParseException("Unterminated quoted field.", 0);
                }

                if (fieldIndex == 0)
                {
                    firstFieldLength = quotedLength;
                }
            }
            else
            {
                if (!TryReadTextUnquotedField(text, delimiter, trim, position, out var field, out var nextPosition))
                {
                    throw new CsvParseException("Unexpected quote in unquoted CSV field.", 0);
                }

                position = nextPosition;
                VisitTextField(field, emitFields, recordIndex, fieldIndex, projectedFieldVisitor, ref fieldVisitor, ref firstFieldLength);
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

            if (!strictQuotes)
            {
                position = recordStart;
                return ReadFlexibleTextRecordFieldSpans(
                    text,
                    delimiter,
                    trim,
                    emitFields,
                    recordIndex,
                    ref position,
                    projectedFieldVisitor,
                    ref fieldVisitor,
                    out firstFieldLength);
            }

            throw new CsvParseException("Unexpected character after CSV field.", 0);
        }

        if (pendingTrailingField)
        {
            VisitTextField(text.Slice(text.Length, 0), emitFields, recordIndex, fieldIndex, projectedFieldVisitor, ref fieldVisitor, ref firstFieldLength);
            fieldIndex++;
        }

        return fieldIndex;
    }

    private static ReadOnlySpan<char> ReadTextUnquotedField(ReadOnlySpan<char> text, char delimiter, bool trim, ref int position)
    {
        if (!TryReadTextUnquotedField(text, delimiter, trim, position, out var field, out var nextPosition))
        {
            throw new CsvParseException("Unexpected quote in unquoted CSV field.", 0);
        }

        position = nextPosition;
        return field;
    }

    private static bool TryReadTextUnquotedField(
        ReadOnlySpan<char> text,
        char delimiter,
        bool trim,
        int position,
        out ReadOnlySpan<char> field,
        out int nextPosition)
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
            nextPosition = text.Length;
        }
        else
        {
            nextPosition = position + terminatorOffset;
            if (text[nextPosition] == '"')
            {
                field = default;
                return false;
            }
        }

        field = text.Slice(start, nextPosition - start);
        if (trim)
        {
            field = TrimTextField(field);
        }

        return true;
    }

    private static bool TryVisitTextQuotedField<TVisitor>(
        ReadOnlySpan<char> text,
        char delimiter,
        bool trim,
        bool emitFields,
        int recordIndex,
        int fieldIndex,
        ref int position,
        ICsvProjectedFieldSpanVisitor? projectedFieldVisitor,
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

            if (position < text.Length &&
                text[position] != delimiter &&
                text[position] != '\r' &&
                text[position] != '\n')
            {
                fieldLength = 0;
                return false;
            }

            fieldLength = valueEnd - valueStart - escapeCount;
            if (!emitFields)
            {
                return true;
            }

            var field = text.Slice(valueStart, valueEnd - valueStart);
            if (!CsvFieldSpanProjection.ShouldVisitField(projectedFieldVisitor, recordIndex, fieldIndex))
            {
                return true;
            }

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
        ICsvProjectedFieldSpanVisitor? projectedFieldVisitor,
        ref TVisitor fieldVisitor,
        ref int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        if (fieldIndex == 0)
        {
            firstFieldLength = value.Length;
        }

        if (emitFields && CsvFieldSpanProjection.ShouldVisitField(projectedFieldVisitor, recordIndex, fieldIndex))
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
        ICsvProjectedFieldSpanVisitor? projectedFieldVisitor,
        ref TVisitor fieldVisitor,
        ref int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        if (trim)
        {
            value = TrimTextField(value);
        }

        VisitTextField(value, emitFields, recordIndex, fieldIndex, projectedFieldVisitor, ref fieldVisitor, ref firstFieldLength);
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

    private static void SkipTextRecord(ReadOnlySpan<char> text, ref int position)
    {
        while (position < text.Length && text[position] != '\r' && text[position] != '\n')
        {
            position++;
        }

        if (position < text.Length)
        {
            ConsumeTextLineSeparator(text, ref position);
        }
    }

    private static bool IsTextW3CFieldsLine(ReadOnlySpan<char> text, int position)
    {
        const string prefix = "#Fields:";
        if (text.Length - position < prefix.Length)
        {
            return false;
        }

        return text.Slice(position, prefix.Length).Equals(prefix.AsSpan(), StringComparison.OrdinalIgnoreCase);
    }
#endif
}
