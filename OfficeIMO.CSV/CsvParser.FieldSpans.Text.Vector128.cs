#nullable enable

#if NET8_0_OR_GREATER
using System.Numerics;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;
#endif

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
#if NET8_0_OR_GREATER
    /// <summary>
    /// Scans the unquoted prefix of a record with portable 128-bit SIMD. Late quoted fields reuse
    /// the delimiter prefix in the shared quoted parsers; early quotes continue in the general
    /// quote-aware SIMD path.
    /// </summary>
    private static bool TryReadTextRecordFieldSpansVector128<TVisitor>(
        ReadOnlySpan<char> text,
        char delimiter,
        bool allowEmpty,
        bool emitFields,
        int recordIndex,
        ref int delimiterIndexCapacity,
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

        const int vectorCharacterCount = 8;
        var start = position;
        var end = text.Length - vectorCharacterCount;
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
        var delimiterVector = Vector128.Create((ushort)delimiter);
        var delimiterCount = 0;
        var pos = start;

        while (pos <= end)
        {
            var values = MemoryMarshal.Cast<char, ushort>(text.Slice(pos, vectorCharacterCount));
            var vector = Vector128.LoadUnsafe(ref MemoryMarshal.GetReference(values));
            var delimiterMask = Vector128.Equals(vector, delimiterVector).ExtractMostSignificantBits();
            var quoteMask = Vector128.Equals(vector, QuoteCharVector128).ExtractMostSignificantBits();
            var carriageReturnMask = Vector128.Equals(vector, CarriageReturnCharVector128).ExtractMostSignificantBits();
            var lineFeedMask = Vector128.Equals(vector, LineFeedCharVector128).ExtractMostSignificantBits();
            var terminalMask = quoteMask | carriageReturnMask | lineFeedMask;

            if (terminalMask != 0)
            {
                var terminalOffset = BitOperations.TrailingZeroCount(terminalMask);
                var delimiterMaskBeforeTerminal = delimiterMask & ((1u << terminalOffset) - 1u);
                if (!AddTextDelimiterIndexes(delimiterMaskBeforeTerminal, pos, delimiterIndexes, ref delimiterCount))
                {
                    delimiterIndexCapacity = 64;
                    return false;
                }

                if (((quoteMask >> terminalOffset) & 1u) != 0)
                {
                    if (delimiterCount < TextQuotedPrefixReuseMinimumDelimiterCount)
                    {
                        return TryReadTextQuoteAwareRecordFieldSpansVector128(
                            text,
                            delimiter,
                            allowEmpty,
                            emitFields,
                            recordIndex,
                            start,
                            delimiterIndexCapacity,
                            ref position,
                            projectedFieldVisitor,
                            ref fieldVisitor,
                            ref scratch,
                            out fieldCount,
                            out firstFieldLength);
                    }

                    var quoteIndex = pos + terminalOffset;
                    if (TryReadTextFinalQuotedRecordFieldSpansFromPrefix(
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
                            out firstFieldLength))
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

            pos += vectorCharacterCount;
        }

        return false;
    }
#endif
}
