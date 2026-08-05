#nullable enable

#if NET8_0_OR_GREATER
using System.Numerics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;

namespace OfficeIMO.CSV;

internal sealed partial class CsvLineReader
{
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    private bool TryReadUnquotedFieldSpansOrLineAvx512<TVisitor>(
        char delimiter,
        bool allowEmpty,
        bool emitFields,
        int recordIndex,
        ICsvProjectedFieldSpanVisitor? projectedFieldVisitor,
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

        int start = _position;
        int end = _length - Vector512<ushort>.Count;
        if (start > end)
        {
            return false;
        }

        Span<int> delimiterIndexes = stackalloc int[UnquotedDelimiterIndexCapacity];
        int delimiterCount = 0;
        int position = start;
        Vector256<byte> delimiterVector = CreateDelimiterVector(delimiter);
        ref ushort bufferStart = ref Unsafe.As<char, ushort>(
            ref MemoryMarshal.GetArrayDataReference(_buffer));

        while (position <= end)
        {
            Vector512<ushort> values = Vector512.LoadUnsafe(ref bufferStart, (nuint)position);
            Vector256<byte> bytes = System.Runtime.Intrinsics.X86.Avx512BW
                .ConvertToVector256ByteWithSaturation(values);
            uint delimiterMask = (uint)System.Runtime.Intrinsics.X86.Avx512BW.MoveMask(
                Vector256.Equals(bytes, delimiterVector));
            uint quoteMask = (uint)System.Runtime.Intrinsics.X86.Avx512BW.MoveMask(
                Vector256.Equals(bytes, CsvQuoteVector));
            uint carriageReturnMask = (uint)System.Runtime.Intrinsics.X86.Avx512BW.MoveMask(
                Vector256.Equals(bytes, CsvCarriageReturnVector));
            uint lineFeedMask = (uint)System.Runtime.Intrinsics.X86.Avx512BW.MoveMask(
                Vector256.Equals(bytes, CsvLineFeedVector));
            uint terminalMask = quoteMask | carriageReturnMask | lineFeedMask;

            if (terminalMask != 0)
            {
                int terminalOffset = BitOperations.TrailingZeroCount(terminalMask);
                uint delimiterMaskBeforeTerminal = delimiterMask & ((1u << terminalOffset) - 1u);
                if (!AddDelimiterIndexes(
                        delimiterMaskBeforeTerminal,
                        position,
                        delimiterIndexes,
                        ref delimiterCount))
                {
                    return false;
                }

                if (((quoteMask >> terminalOffset) & 1u) != 0)
                {
                    int quoteIndex = position + terminalOffset;
                    if (delimiterCount >= QuotedPrefixReuseMinimumDelimiterCount &&
                        TryReadStandardQuotedFieldSpansOrLineFromPrefix(
                            delimiter,
                            allowEmpty,
                            emitFields,
                            recordIndex,
                            delimiterIndexes.Slice(0, delimiterCount),
                            quoteIndex,
                            projectedFieldVisitor,
                            ref fieldVisitor,
                            out fieldCount,
                            out isEmptyRecord,
                            out separator,
                            out readResult))
                    {
                        return true;
                    }

                    return false;
                }

                int newlineIndex = position + terminalOffset;
                int lineLength = newlineIndex - start;
                fieldCount = VisitIndexedUntrimmedUnquotedFieldSpans(
                    start,
                    newlineIndex,
                    delimiterIndexes.Slice(0, delimiterCount),
                    (allowEmpty || lineLength != 0) && emitFields,
                    recordIndex,
                    projectedFieldVisitor,
                    ref fieldVisitor,
                    out int firstFieldLength);
                isEmptyRecord = fieldCount == 1 && firstFieldLength == 0;
                _position = newlineIndex;
                ConsumeLineSeparator(_buffer[newlineIndex], out separator);
                readResult = CsvLineReadResult.UnquotedRecord;
                return true;
            }

            if (!AddDelimiterIndexes(delimiterMask, position, delimiterIndexes, ref delimiterCount))
            {
                return false;
            }

            position += Vector512<ushort>.Count;
        }

        return false;
    }
}
#endif
