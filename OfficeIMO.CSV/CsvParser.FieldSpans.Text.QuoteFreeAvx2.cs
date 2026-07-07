namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
#if NET8_0_OR_GREATER
    private static bool TryReadTextQuoteFreeRecordFieldSpansAvx2<TVisitor>(
        ReadOnlySpan<char> text,
        char delimiter,
        bool allowEmpty,
        bool emitFields,
        int recordIndex,
        System.Runtime.Intrinsics.Vector256<byte> delimiterVector,
        ref int position,
        ref TVisitor fieldVisitor,
        out int fieldCount,
        out int firstFieldLength)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        fieldCount = 0;
        firstFieldLength = 0;

        var start = position;
        var end = text.Length - 32;
        if (start > end)
        {
            return false;
        }

        var fieldStart = start;
        var fieldIndex = 0;
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
            var carriageReturnMask = (uint)System.Runtime.Intrinsics.X86.Avx2.MoveMask(
                System.Runtime.Intrinsics.X86.Avx2.CompareEqual(packedBytes, CarriageReturnByteVector));
            var lineFeedMask = (uint)System.Runtime.Intrinsics.X86.Avx2.MoveMask(
                System.Runtime.Intrinsics.X86.Avx2.CompareEqual(packedBytes, LineFeedByteVector));
            var terminalMask = carriageReturnMask | lineFeedMask;
            var specialMask = delimiterMask | terminalMask;

            if (specialMask != 0)
            {
                while (specialMask != 0)
                {
                    var offset = System.Numerics.BitOperations.TrailingZeroCount(specialMask);
                    var bit = 1u << offset;
                    specialMask &= specialMask - 1;
                    var absoluteIndex = pos + offset;

                    if ((bit & delimiterMask) != 0)
                    {
                        var length = absoluteIndex - fieldStart;
                        if (fieldIndex == 0)
                        {
                            firstFieldLength = length;
                        }

                        if (emitFields)
                        {
                            fieldVisitor.VisitField(recordIndex, fieldIndex, text.Slice(fieldStart, length));
                        }

                        fieldIndex++;
                        fieldStart = absoluteIndex + 1;
                        continue;
                    }

                    var lineLength = absoluteIndex - start;
                    var emitRecordFields = (allowEmpty || lineLength != 0) && emitFields;
                    var finalLength = absoluteIndex - fieldStart;
                    if (fieldIndex == 0)
                    {
                        firstFieldLength = finalLength;
                    }

                    if (emitRecordFields)
                    {
                        fieldVisitor.VisitField(recordIndex, fieldIndex, text.Slice(fieldStart, finalLength));
                    }

                    fieldCount = fieldIndex + 1;
                    position = absoluteIndex;
                    ConsumeTextLineSeparator(text, ref position);
                    return true;
                }
            }

            pos += 32;
        }

        while (pos < text.Length)
        {
            var value = text[pos];
            if (value == delimiter)
            {
                var length = pos - fieldStart;
                if (fieldIndex == 0)
                {
                    firstFieldLength = length;
                }

                if (emitFields)
                {
                    fieldVisitor.VisitField(recordIndex, fieldIndex, text.Slice(fieldStart, length));
                }

                fieldIndex++;
                fieldStart = pos + 1;
                pos++;
                continue;
            }

            if (value == '\r' || value == '\n')
            {
                var lineLength = pos - start;
                var emitRecordFields = (allowEmpty || lineLength != 0) && emitFields;
                var finalLength = pos - fieldStart;
                if (fieldIndex == 0)
                {
                    firstFieldLength = finalLength;
                }

                if (emitRecordFields)
                {
                    fieldVisitor.VisitField(recordIndex, fieldIndex, text.Slice(fieldStart, finalLength));
                }

                fieldCount = fieldIndex + 1;
                position = pos;
                ConsumeTextLineSeparator(text, ref position);
                return true;
            }

            pos++;
        }

        if (pos == start)
        {
            return false;
        }

        var tailLength = text.Length - fieldStart;
        if (fieldIndex == 0)
        {
            firstFieldLength = tailLength;
        }

        if ((allowEmpty || text.Length != start) && emitFields)
        {
            fieldVisitor.VisitField(recordIndex, fieldIndex, text.Slice(fieldStart, tailLength));
        }

        fieldCount = fieldIndex + 1;
        position = text.Length;
        return true;
    }
#endif
}
