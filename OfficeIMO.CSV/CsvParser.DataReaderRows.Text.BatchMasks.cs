#if NET8_0_OR_GREATER
using System.Runtime.CompilerServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.Arm;
using System.Runtime.Intrinsics.X86;

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
    // Delimiters are restricted to 1..254. Saturation maps UTF-16 outside that
    // range to 0 or 255, so its low byte cannot become CSV punctuation.
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static uint GetTextDataReaderBatchMask(ref ushort text, int position, byte delimiter, bool textMayContainQuote)
    {
        if (Avx2.IsSupported)
        {
            Vector256<byte> packedBytes;
            if (Avx512BW.IsSupported)
            {
                packedBytes = Avx512BW.ConvertToVector256ByteWithSaturation(
                    Vector512.LoadUnsafe(ref text, (nuint)position));
            }
            else
            {
                var first = Vector256.LoadUnsafe(ref text, (nuint)position).AsInt16();
                var second = Vector256.LoadUnsafe(ref text, (nuint)(position + 16)).AsInt16();
                var packed = Avx2.PackUnsignedSaturate(first, second);
                packedBytes = Avx2.Permute4x64(packed.AsInt64(), 0b11_01_10_00).AsByte();
            }

            var matches = Avx2.Or(Avx2.CompareEqual(packedBytes, Vector256.Create(delimiter)),
                Avx2.Or(Avx2.CompareEqual(packedBytes, CarriageReturnByteVector),
                    Avx2.CompareEqual(packedBytes, LineFeedByteVector)));
            if (textMayContainQuote) matches = Avx2.Or(matches, Avx2.CompareEqual(packedBytes, QuoteByteVector));
            return (uint)Avx2.MoveMask(matches);
        }

        var lower = GetTextDataReaderBatchArmMatches(PackTextDataReaderBatchArm(ref text, position), delimiter, textMayContainQuote);
        var upper = GetTextDataReaderBatchArmMatches(PackTextDataReaderBatchArm(ref text, position + 16), delimiter, textMayContainQuote);
        return Vector128.ExtractMostSignificantBits(lower) | (Vector128.ExtractMostSignificantBits(upper) << 16);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static Vector128<byte> PackTextDataReaderBatchArm(ref ushort text, int position) =>
        AdvSimd.ExtractNarrowingSaturateUnsignedUpper(
            AdvSimd.ExtractNarrowingSaturateUnsignedLower(
                Vector128.LoadUnsafe(ref text, (nuint)position).AsInt16()),
            Vector128.LoadUnsafe(ref text, (nuint)(position + 8)).AsInt16());

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static Vector128<byte> GetTextDataReaderBatchArmMatches(Vector128<byte> bytes, byte delimiter, bool textMayContainQuote)
    {
        var matches = Vector128.Equals(bytes, Vector128.Create(delimiter)) |
            Vector128.Equals(bytes, Vector128.Create((byte)'\r')) |
            Vector128.Equals(bytes, Vector128.Create((byte)'\n'));
        if (textMayContainQuote) matches |= Vector128.Equals(bytes, Vector128.Create((byte)'"'));
        return matches;
    }
}
#endif
