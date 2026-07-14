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
    private static readonly Vector128<ushort> CarriageReturnCharVector128 = Vector128.Create((ushort)'\r');
    private static readonly Vector128<ushort> LineFeedCharVector128 = Vector128.Create((ushort)'\n');

    /// <summary>
    /// Reads an unquoted record with portable 128-bit SIMD. This complements the wider AVX2 path on
    /// Arm64 and other runtimes where 128-bit vectors are hardware accelerated.
    /// </summary>
    private static bool TryReadTextQuoteFreeRecordFieldSpansVector128<TVisitor>(
        ReadOnlySpan<char> text,
        char delimiter,
        bool allowEmpty,
        bool emitFields,
        int recordIndex,
        ref int position,
        ICsvProjectedFieldSpanVisitor? projectedFieldVisitor,
        ref TVisitor fieldVisitor,
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

        var delimiterVector = Vector128.Create((ushort)delimiter);
        var fieldStart = start;
        var fieldIndex = 0;
        var pos = start;

        while (pos <= end)
        {
            var values = MemoryMarshal.Cast<char, ushort>(text.Slice(pos, vectorCharacterCount));
            var vector = Vector128.LoadUnsafe(ref MemoryMarshal.GetReference(values));
            var delimiterMask = Vector128.Equals(vector, delimiterVector).ExtractMostSignificantBits();
            var carriageReturnMask = Vector128.Equals(vector, CarriageReturnCharVector128).ExtractMostSignificantBits();
            var lineFeedMask = Vector128.Equals(vector, LineFeedCharVector128).ExtractMostSignificantBits();
            var terminalMask = carriageReturnMask | lineFeedMask;
            var specialMask = delimiterMask | terminalMask;

            while (specialMask != 0)
            {
                var offset = BitOperations.TrailingZeroCount(specialMask);
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

                    if (emitFields && CsvFieldSpanProjection.ShouldVisitField(projectedFieldVisitor, recordIndex, fieldIndex))
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

                if (emitRecordFields && CsvFieldSpanProjection.ShouldVisitField(projectedFieldVisitor, recordIndex, fieldIndex))
                {
                    fieldVisitor.VisitField(recordIndex, fieldIndex, text.Slice(fieldStart, finalLength));
                }

                fieldCount = fieldIndex + 1;
                position = absoluteIndex;
                ConsumeTextLineSeparator(text, ref position);
                return true;
            }

            pos += vectorCharacterCount;
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

                if (emitFields && CsvFieldSpanProjection.ShouldVisitField(projectedFieldVisitor, recordIndex, fieldIndex))
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

                if (emitRecordFields && CsvFieldSpanProjection.ShouldVisitField(projectedFieldVisitor, recordIndex, fieldIndex))
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

        if ((allowEmpty || text.Length != start) && emitFields && CsvFieldSpanProjection.ShouldVisitField(projectedFieldVisitor, recordIndex, fieldIndex))
        {
            fieldVisitor.VisitField(recordIndex, fieldIndex, text.Slice(fieldStart, tailLength));
        }

        fieldCount = fieldIndex + 1;
        position = text.Length;
        return true;
    }
#endif
}
