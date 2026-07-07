#nullable enable

using System.Numerics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
#if NET8_0_OR_GREATER
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;
#endif

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
#if NET8_0_OR_GREATER
    private const int TextQuoteAwareFieldSpanCapacity = 64;

    private static bool TryReadTextQuoteAwareRecordFieldSpansAvx2<TVisitor>(
        ReadOnlySpan<char> text,
        char delimiter,
        bool allowEmpty,
        bool emitFields,
        int recordIndex,
        int recordStart,
        ref int position,
        ref TVisitor fieldVisitor,
        ref char[]? scratch,
        out int fieldCount,
        out int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        fieldCount = 0;
        firstFieldLength = 0;
        if (!Avx2.IsSupported || delimiter > byte.MaxValue)
        {
            return false;
        }

        Span<TextQuoteAwareFieldSpan> fields = stackalloc TextQuoteAwareFieldSpan[TextQuoteAwareFieldSpanCapacity];
        var fieldStart = recordStart;
        var quoteCount = 0;
        var delimiterVector = Vector256.Create((byte)delimiter);
        return ContinueTextQuoteAwareRecordFieldSpansAvx2(
            text,
            delimiter,
            delimiterVector,
            allowEmpty,
            emitFields,
            recordIndex,
            recordStart,
            recordStart,
            ref fieldStart,
            ref quoteCount,
            fields,
            ref fieldCount,
            ref position,
            ref fieldVisitor,
            ref scratch,
            out firstFieldLength);
    }

    private static bool TryReadTextQuoteAwareRecordFieldSpansAvx2FromCurrentChunk<TVisitor>(
        ReadOnlySpan<char> text,
        char delimiter,
        bool allowEmpty,
        bool emitFields,
        int recordIndex,
        int recordStart,
        int chunkStart,
        uint delimiterMask,
        uint quoteMask,
        uint carriageReturnMask,
        uint lineFeedMask,
        Vector256<byte> delimiterVector,
        ref int position,
        ref TVisitor fieldVisitor,
        ref char[]? scratch,
        out int fieldCount,
        out int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        fieldCount = 0;
        firstFieldLength = 0;
        if (!Avx2.IsSupported || delimiter > byte.MaxValue)
        {
            return false;
        }

        Span<TextQuoteAwareFieldSpan> fields = stackalloc TextQuoteAwareFieldSpan[TextQuoteAwareFieldSpanCapacity];
        var fieldStart = recordStart;
        var quoteCount = 0;
        var specialMask = delimiterMask | quoteMask | carriageReturnMask | lineFeedMask;
        if (!ProcessTextQuoteAwareMask(
            text,
            specialMask,
            delimiterMask,
            quoteMask,
            carriageReturnMask,
            lineFeedMask,
            chunkStart,
            ref fieldStart,
            ref quoteCount,
            fields,
            ref fieldCount,
            out var recordEnd,
            out var nextPosition))
        {
            return false;
        }

        if (recordEnd >= 0)
        {
            return CompleteTextQuoteAwareRecord(
                text,
                fields.Slice(0, fieldCount),
                allowEmpty,
                emitFields,
                recordIndex,
                recordStart,
                nextPosition,
                ref position,
                ref fieldVisitor,
                ref scratch,
                out firstFieldLength);
        }

        return ContinueTextQuoteAwareRecordFieldSpansAvx2(
            text,
            delimiter,
            delimiterVector,
            allowEmpty,
            emitFields,
            recordIndex,
            recordStart,
            chunkStart + 32,
            ref fieldStart,
            ref quoteCount,
            fields,
            ref fieldCount,
            ref position,
            ref fieldVisitor,
            ref scratch,
            out firstFieldLength);
    }

    private static bool ContinueTextQuoteAwareRecordFieldSpansAvx2<TVisitor>(
        ReadOnlySpan<char> text,
        char delimiter,
        Vector256<byte> delimiterVector,
        bool allowEmpty,
        bool emitFields,
        int recordIndex,
        int recordStart,
        int index,
        ref int fieldStart,
        ref int quoteCount,
        Span<TextQuoteAwareFieldSpan> fields,
        ref int fieldCount,
        ref int position,
        ref TVisitor fieldVisitor,
        ref char[]? scratch,
        out int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        firstFieldLength = 0;
        var end = text.Length - 32;

        while (index <= end)
        {
            var values = MemoryMarshal.Cast<char, short>(text.Slice(index, 32));
            var first = Vector256.LoadUnsafe(ref MemoryMarshal.GetReference(values));
            var second = Vector256.LoadUnsafe(ref MemoryMarshal.GetReference(values.Slice(16)));
            var packed = Avx2.PackUnsignedSaturate(first, second);
            var packedBytes = Vector256.AsByte(
                Avx2.Permute4x64(Vector256.AsInt64(packed), 0b11_01_10_00));

            var delimiterMask = (uint)Avx2.MoveMask(Avx2.CompareEqual(packedBytes, delimiterVector));
            var quoteMask = (uint)Avx2.MoveMask(Avx2.CompareEqual(packedBytes, QuoteByteVector));
            var carriageReturnMask = (uint)Avx2.MoveMask(Avx2.CompareEqual(packedBytes, CarriageReturnByteVector));
            var lineFeedMask = (uint)Avx2.MoveMask(Avx2.CompareEqual(packedBytes, LineFeedByteVector));
            var specialMask = delimiterMask | quoteMask | carriageReturnMask | lineFeedMask;

            if (specialMask != 0)
            {
                if (!ProcessTextQuoteAwareMask(
                    text,
                    specialMask,
                    delimiterMask,
                    quoteMask,
                    carriageReturnMask,
                    lineFeedMask,
                    index,
                    ref fieldStart,
                    ref quoteCount,
                    fields,
                    ref fieldCount,
                    out var recordEnd,
                    out var nextPosition))
                {
                    return false;
                }

                if (recordEnd >= 0)
                {
                    return CompleteTextQuoteAwareRecord(
                        text,
                        fields.Slice(0, fieldCount),
                        allowEmpty,
                        emitFields,
                        recordIndex,
                        recordStart,
                        nextPosition,
                        ref position,
                        ref fieldVisitor,
                        ref scratch,
                        out firstFieldLength);
                }
            }

            index += 32;
        }

        if (!ProcessTextQuoteAwareTail(
            text,
            delimiter,
            index,
            ref fieldStart,
            ref quoteCount,
            fields,
            ref fieldCount,
            out var tailNextPosition))
        {
            return false;
        }

        return CompleteTextQuoteAwareRecord(
            text,
            fields.Slice(0, fieldCount),
            allowEmpty,
            emitFields,
            recordIndex,
            recordStart,
            tailNextPosition,
            ref position,
            ref fieldVisitor,
            ref scratch,
            out firstFieldLength);
    }

    private static bool ProcessTextQuoteAwareMask(
        ReadOnlySpan<char> text,
        uint specialMask,
        uint delimiterMask,
        uint quoteMask,
        uint carriageReturnMask,
        uint lineFeedMask,
        int chunkStart,
        ref int fieldStart,
        ref int quoteCount,
        Span<TextQuoteAwareFieldSpan> fields,
        ref int fieldCount,
        out int recordEnd,
        out int nextPosition)
    {
        recordEnd = -1;
        nextPosition = -1;

        while (specialMask != 0)
        {
            var offset = BitOperations.TrailingZeroCount(specialMask);
            var bit = 1u << offset;
            specialMask &= specialMask - 1;

            if ((bit & quoteMask) != 0)
            {
                quoteCount++;
                continue;
            }

            if ((quoteCount & 1) != 0)
            {
                if ((bit & (carriageReturnMask | lineFeedMask)) != 0)
                {
                    return false;
                }

                continue;
            }

            var absoluteIndex = chunkStart + offset;
            if ((bit & delimiterMask) != 0)
            {
                if (!TryAddTextQuoteAwareField(fields, ref fieldCount, fieldStart, absoluteIndex, quoteCount))
                {
                    return false;
                }

                fieldStart = absoluteIndex + 1;
                quoteCount = 0;
                continue;
            }

            if ((bit & (carriageReturnMask | lineFeedMask)) != 0)
            {
                if (!TryAddTextQuoteAwareField(fields, ref fieldCount, fieldStart, absoluteIndex, quoteCount))
                {
                    return false;
                }

                recordEnd = absoluteIndex;
                nextPosition = absoluteIndex + 1;
                if ((bit & carriageReturnMask) != 0 &&
                    nextPosition < text.Length &&
                    text[nextPosition] == '\n')
                {
                    nextPosition++;
                }

                return true;
            }
        }

        return true;
    }

    private static bool ProcessTextQuoteAwareTail(
        ReadOnlySpan<char> text,
        char delimiter,
        int index,
        ref int fieldStart,
        ref int quoteCount,
        Span<TextQuoteAwareFieldSpan> fields,
        ref int fieldCount,
        out int nextPosition)
    {
        while (index < text.Length)
        {
            var value = text[index];
            if (value == '"')
            {
                quoteCount++;
                index++;
                continue;
            }

            if ((quoteCount & 1) == 0)
            {
                if (value == delimiter)
                {
                    if (!TryAddTextQuoteAwareField(fields, ref fieldCount, fieldStart, index, quoteCount))
                    {
                        nextPosition = -1;
                        return false;
                    }

                    fieldStart = index + 1;
                    quoteCount = 0;
                }
                else if (value == '\r' || value == '\n')
                {
                    if (!TryAddTextQuoteAwareField(fields, ref fieldCount, fieldStart, index, quoteCount))
                    {
                        nextPosition = -1;
                        return false;
                    }

                    nextPosition = index + 1;
                    if (value == '\r' && nextPosition < text.Length && text[nextPosition] == '\n')
                    {
                        nextPosition++;
                    }

                    return true;
                }
            }
            else if (value == '\r' || value == '\n')
            {
                nextPosition = -1;
                return false;
            }

            index++;
        }

        if (quoteCount != 0 && (quoteCount & 1) != 0)
        {
            nextPosition = -1;
            return false;
        }

        if (!TryAddTextQuoteAwareField(fields, ref fieldCount, fieldStart, text.Length, quoteCount))
        {
            nextPosition = -1;
            return false;
        }

        nextPosition = text.Length;
        return true;
    }

    private static bool CompleteTextQuoteAwareRecord<TVisitor>(
        ReadOnlySpan<char> text,
        ReadOnlySpan<TextQuoteAwareFieldSpan> fields,
        bool allowEmpty,
        bool emitFields,
        int recordIndex,
        int recordStart,
        int nextPosition,
        ref int position,
        ref TVisitor fieldVisitor,
        ref char[]? scratch,
        out int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        firstFieldLength = 0;
        if (fields.Length == 0)
        {
            return false;
        }

        if (emitFields && (allowEmpty || fields[^1].End > recordStart))
        {
            if (!VisitTextQuoteAwareFieldSpans(text, fields, recordIndex, ref fieldVisitor, ref scratch, out firstFieldLength))
            {
                return false;
            }
        }
        else
        {
            if (!TryGetTextQuoteAwareFirstFieldLength(text, fields[0], out firstFieldLength))
            {
                return false;
            }
        }

        position = nextPosition;
        return true;
    }

    private static bool VisitTextQuoteAwareFieldSpans<TVisitor>(
        ReadOnlySpan<char> text,
        ReadOnlySpan<TextQuoteAwareFieldSpan> fields,
        int recordIndex,
        ref TVisitor fieldVisitor,
        ref char[]? scratch,
        out int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        firstFieldLength = 0;
        for (var fieldIndex = 0; fieldIndex < fields.Length; fieldIndex++)
        {
            var field = fields[fieldIndex];
            if (!TryVisitTextQuoteAwareField(text, field, recordIndex, fieldIndex, ref fieldVisitor, ref scratch, out var fieldLength))
            {
                return false;
            }

            if (fieldIndex == 0)
            {
                firstFieldLength = fieldLength;
            }
        }

        return true;
    }

    private static bool TryVisitTextQuoteAwareField<TVisitor>(
        ReadOnlySpan<char> text,
        TextQuoteAwareFieldSpan field,
        int recordIndex,
        int fieldIndex,
        ref TVisitor fieldVisitor,
        ref char[]? scratch,
        out int fieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        if (field.QuoteCount == 0)
        {
            fieldLength = field.End - field.Start;
            fieldVisitor.VisitField(recordIndex, fieldIndex, text.Slice(field.Start, fieldLength));
            return true;
        }

        return TryVisitTextQuoteAwareQuotedField(text, field, recordIndex, fieldIndex, ref fieldVisitor, ref scratch, out fieldLength);
    }

    private static bool TryVisitTextQuoteAwareQuotedField<TVisitor>(
        ReadOnlySpan<char> text,
        TextQuoteAwareFieldSpan field,
        int recordIndex,
        int fieldIndex,
        ref TVisitor fieldVisitor,
        ref char[]? scratch,
        out int fieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        var rawLength = field.End - field.Start;
        if ((field.QuoteCount & 1) != 0 ||
            rawLength < 2 ||
            text[field.Start] != '"' ||
            text[field.End - 1] != '"')
        {
            fieldLength = 0;
            return false;
        }

        var escapedQuoteCount = (field.QuoteCount - 2) / 2;
        fieldLength = rawLength - 2 - escapedQuoteCount;
        var value = text.Slice(field.Start + 1, rawLength - 2);
        if (escapedQuoteCount == 0)
        {
            fieldVisitor.VisitField(recordIndex, fieldIndex, value);
            return true;
        }

        if (fieldVisitor.TryVisitEscapedField(recordIndex, fieldIndex, value, fieldLength))
        {
            return true;
        }

        var firstEscapedQuote = value.IndexOf('"');
        if (firstEscapedQuote < 0)
        {
            return false;
        }

        var unescaped = UnescapeTextQuotedField(value, firstEscapedQuote, fieldLength, ref scratch);
        fieldVisitor.VisitField(recordIndex, fieldIndex, unescaped);
        return true;
    }

    private static bool TryGetTextQuoteAwareFirstFieldLength(
        ReadOnlySpan<char> text,
        TextQuoteAwareFieldSpan field,
        out int fieldLength)
    {
        if (field.QuoteCount == 0)
        {
            fieldLength = field.End - field.Start;
            return true;
        }

        var rawLength = field.End - field.Start;
        if ((field.QuoteCount & 1) != 0 ||
            rawLength < 2 ||
            text[field.Start] != '"' ||
            text[field.End - 1] != '"')
        {
            fieldLength = 0;
            return false;
        }

        fieldLength = rawLength - 2 - ((field.QuoteCount - 2) / 2);
        return true;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static bool TryAddTextQuoteAwareField(
        Span<TextQuoteAwareFieldSpan> fields,
        ref int fieldCount,
        int start,
        int end,
        int quoteCount)
    {
        if (fieldCount >= fields.Length)
        {
            return false;
        }

        fields[fieldCount++] = new TextQuoteAwareFieldSpan(start, end, quoteCount);
        return true;
    }

    private readonly struct TextQuoteAwareFieldSpan
    {
        public TextQuoteAwareFieldSpan(int start, int end, int quoteCount)
        {
            Start = start;
            End = end;
            QuoteCount = quoteCount;
        }

        public int Start { get; }

        public int End { get; }

        public int QuoteCount { get; }
    }
#endif
}
