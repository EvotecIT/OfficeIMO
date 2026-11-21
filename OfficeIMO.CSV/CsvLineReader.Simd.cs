#if NET8_0_OR_GREATER
namespace OfficeIMO.CSV;

internal sealed partial class CsvLineReader
{
    private bool TryReadUnquotedRecordOrLineSimd(
        char delimiter,
        List<string> fields,
        out string? line,
        out string separator,
        out CsvLineReadResult readResult)
    {
        line = null;
        separator = string.Empty;
        readResult = CsvLineReadResult.Line;

        var start = _position;
        var end = _length - 32;
        if (start > end)
        {
            return false;
        }

        Span<int> delimiterIndexes = stackalloc int[UnquotedDelimiterIndexCapacity];
        var delimiterCount = 0;
        var pos = start;
        var delimiterVector = CreateDelimiterVector(delimiter);

        while (pos <= end)
        {
            uint delimiterMask, quoteMask, newlineMask;
            if (System.Runtime.Intrinsics.X86.Avx2.IsSupported)
            {
                var values = System.Runtime.InteropServices.MemoryMarshal.Cast<char, short>(_buffer.AsSpan(pos, 32));
                var first = System.Runtime.Intrinsics.Vector256.LoadUnsafe(ref System.Runtime.InteropServices.MemoryMarshal.GetReference(values));
                var second = System.Runtime.Intrinsics.Vector256.LoadUnsafe(ref System.Runtime.InteropServices.MemoryMarshal.GetReference(values.Slice(16)));
                var packed = System.Runtime.Intrinsics.X86.Avx2.PackUnsignedSaturate(first, second);
                var packedBytes = System.Runtime.Intrinsics.Vector256.AsByte(
                    System.Runtime.Intrinsics.X86.Avx2.Permute4x64(System.Runtime.Intrinsics.Vector256.AsInt64(packed), 0b11_01_10_00));

                delimiterMask = (uint)System.Runtime.Intrinsics.X86.Avx2.MoveMask(
                    System.Runtime.Intrinsics.X86.Avx2.CompareEqual(packedBytes, delimiterVector));
                quoteMask = (uint)System.Runtime.Intrinsics.X86.Avx2.MoveMask(
                    System.Runtime.Intrinsics.X86.Avx2.CompareEqual(packedBytes, CsvQuoteVector));
                var carriageReturnMask = (uint)System.Runtime.Intrinsics.X86.Avx2.MoveMask(
                    System.Runtime.Intrinsics.X86.Avx2.CompareEqual(packedBytes, CsvCarriageReturnVector));
                var lineFeedMask = (uint)System.Runtime.Intrinsics.X86.Avx2.MoveMask(
                    System.Runtime.Intrinsics.X86.Avx2.CompareEqual(packedBytes, CsvLineFeedVector));
                newlineMask = carriageReturnMask | lineFeedMask;
            }
            else
            {
                ref ushort text = ref System.Runtime.CompilerServices.Unsafe.As<char, ushort>(
                    ref System.Runtime.InteropServices.MemoryMarshal.GetArrayDataReference(_buffer));
                CsvSimdMasks.ReadText32Arm(ref text, pos, (byte)delimiter,
                    out delimiterMask, out quoteMask, out newlineMask);
            }
            quoteMask &= _quoteMask;
            var terminalMask = quoteMask | newlineMask;

            if (terminalMask != 0)
            {
                var terminalOffset = System.Numerics.BitOperations.TrailingZeroCount(terminalMask);
                var delimiterMaskBeforeTerminal = delimiterMask & ((1u << terminalOffset) - 1u);
                if (!AddDelimiterIndexes(delimiterMaskBeforeTerminal, pos, delimiterIndexes, ref delimiterCount))
                {
                    return false;
                }

                if (((quoteMask >> terminalOffset) & 1u) != 0)
                {
                    return false;
                }

                var newlineIndex = pos + terminalOffset;
                AddIndexedUnquotedFields(start, newlineIndex, delimiterIndexes.Slice(0, delimiterCount), fields);
                _position = newlineIndex;
                ConsumeLineSeparator(_buffer[newlineIndex], out separator);
                readResult = CsvLineReadResult.UnquotedRecord;
                return true;
            }

            if (!AddDelimiterIndexes(delimiterMask, pos, delimiterIndexes, ref delimiterCount))
            {
                return false;
            }

            pos += 32;
        }

        return false;
    }

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveOptimization)]
    private bool TryReadUnquotedFieldSpansOrLineSimd<TVisitor>(
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

        var start = _position;
        var end = _length - 32;
        if (start > end)
        {
            return false;
        }

        Span<int> delimiterIndexes = stackalloc int[UnquotedDelimiterIndexCapacity];
        var delimiterCount = 0;
        var pos = start;
        var delimiterVector = CreateDelimiterVector(delimiter);

        while (pos <= end)
        {
            uint delimiterMask, quoteMask, newlineMask;
            if (System.Runtime.Intrinsics.X86.Avx2.IsSupported)
            {
                var values = System.Runtime.InteropServices.MemoryMarshal.Cast<char, short>(_buffer.AsSpan(pos, 32));
                var first = System.Runtime.Intrinsics.Vector256.LoadUnsafe(ref System.Runtime.InteropServices.MemoryMarshal.GetReference(values));
                var second = System.Runtime.Intrinsics.Vector256.LoadUnsafe(ref System.Runtime.InteropServices.MemoryMarshal.GetReference(values.Slice(16)));
                var packed = System.Runtime.Intrinsics.X86.Avx2.PackUnsignedSaturate(first, second);
                var packedBytes = System.Runtime.Intrinsics.Vector256.AsByte(
                    System.Runtime.Intrinsics.X86.Avx2.Permute4x64(System.Runtime.Intrinsics.Vector256.AsInt64(packed), 0b11_01_10_00));

                delimiterMask = (uint)System.Runtime.Intrinsics.X86.Avx2.MoveMask(
                    System.Runtime.Intrinsics.X86.Avx2.CompareEqual(packedBytes, delimiterVector));
                quoteMask = (uint)System.Runtime.Intrinsics.X86.Avx2.MoveMask(
                    System.Runtime.Intrinsics.X86.Avx2.CompareEqual(packedBytes, CsvQuoteVector));
                var carriageReturnMask = (uint)System.Runtime.Intrinsics.X86.Avx2.MoveMask(
                    System.Runtime.Intrinsics.X86.Avx2.CompareEqual(packedBytes, CsvCarriageReturnVector));
                var lineFeedMask = (uint)System.Runtime.Intrinsics.X86.Avx2.MoveMask(
                    System.Runtime.Intrinsics.X86.Avx2.CompareEqual(packedBytes, CsvLineFeedVector));
                newlineMask = carriageReturnMask | lineFeedMask;
            }
            else
            {
                ref ushort text = ref System.Runtime.CompilerServices.Unsafe.As<char, ushort>(
                    ref System.Runtime.InteropServices.MemoryMarshal.GetArrayDataReference(_buffer));
                CsvSimdMasks.ReadText32Arm(ref text, pos, (byte)delimiter,
                    out delimiterMask, out quoteMask, out newlineMask);
            }
            quoteMask &= _quoteMask;
            var terminalMask = quoteMask | newlineMask;

            if (terminalMask != 0)
            {
                var terminalOffset = System.Numerics.BitOperations.TrailingZeroCount(terminalMask);
                var delimiterMaskBeforeTerminal = delimiterMask & ((1u << terminalOffset) - 1u);
                if (!AddDelimiterIndexes(delimiterMaskBeforeTerminal, pos, delimiterIndexes, ref delimiterCount))
                {
                    return false;
                }

                if (((quoteMask >> terminalOffset) & 1u) != 0)
                {
                    var quoteIndex = pos + terminalOffset;
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

                var newlineIndex = pos + terminalOffset;
                var lineLength = newlineIndex - start;
                fieldCount = VisitIndexedUntrimmedUnquotedFieldSpans(
                    start,
                    newlineIndex,
                    delimiterIndexes.Slice(0, delimiterCount),
                    (allowEmpty || lineLength != 0) && emitFields,
                    recordIndex,
                    projectedFieldVisitor,
                    ref fieldVisitor,
                    out var firstFieldLength);
                isEmptyRecord = fieldCount == 1 && firstFieldLength == 0;
                _position = newlineIndex;
                ConsumeLineSeparator(_buffer[newlineIndex], out separator, deferLineFeedRead: true);
                readResult = CsvLineReadResult.UnquotedRecord;
                return true;
            }

            if (!AddDelimiterIndexes(delimiterMask, pos, delimiterIndexes, ref delimiterCount))
            {
                return false;
            }

            pos += 32;
        }

        return false;
    }
}
#endif
