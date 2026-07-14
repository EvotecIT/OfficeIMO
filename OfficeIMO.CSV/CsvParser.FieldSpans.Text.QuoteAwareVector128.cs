#nullable enable

#if NET8_0_OR_GREATER
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;
#endif

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
#if NET8_0_OR_GREATER
    private static readonly Vector128<ushort> QuoteCharVector128 = Vector128.Create((ushort)'"');

    /// <summary>
    /// Reads a standard quoted or multiline record with portable 128-bit SIMD on runtimes that do
    /// not expose AVX2. Validation falls back to the general parser when the fast-path contract is
    /// not satisfied.
    /// </summary>
    private static bool TryReadTextQuoteAwareRecordFieldSpansVector128<TVisitor>(
        ReadOnlySpan<char> text,
        char delimiter,
        bool allowEmpty,
        bool emitFields,
        int recordIndex,
        int recordStart,
        int fieldSpanCapacity,
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
        if (!Vector128.IsHardwareAccelerated)
        {
            return false;
        }

        Span<TextQuoteAwareFieldSpan> fields = fieldSpanCapacity switch
        {
            16 => stackalloc TextQuoteAwareFieldSpan[16],
            32 => stackalloc TextQuoteAwareFieldSpan[32],
            _ => stackalloc TextQuoteAwareFieldSpan[TextQuoteAwareFieldSpanCapacity],
        };
        var delimiterVector = Vector128.Create((ushort)delimiter);
        var fieldStart = recordStart;
        var quoteCount = 0;
        var index = recordStart;
        var end = text.Length - 8;

        while (index <= end)
        {
            var values = MemoryMarshal.Cast<char, ushort>(text.Slice(index, 8));
            var vector = Vector128.LoadUnsafe(ref MemoryMarshal.GetReference(values));
            var delimiterMask = Vector128.Equals(vector, delimiterVector).ExtractMostSignificantBits();
            var quoteMask = Vector128.Equals(vector, QuoteCharVector128).ExtractMostSignificantBits();
            var carriageReturnMask = Vector128.Equals(vector, CarriageReturnCharVector128).ExtractMostSignificantBits();
            var lineFeedMask = Vector128.Equals(vector, LineFeedCharVector128).ExtractMostSignificantBits();
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
                        projectedFieldVisitor,
                        ref fieldVisitor,
                        ref scratch,
                        out firstFieldLength);
                }
            }

            index += 8;
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
            projectedFieldVisitor,
            ref fieldVisitor,
            ref scratch,
            out firstFieldLength);
    }

#endif
}
