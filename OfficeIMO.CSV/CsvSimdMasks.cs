#if NET8_0_OR_GREATER
using System.Runtime.CompilerServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.Arm;

namespace OfficeIMO.CSV;

/// <summary>Block masks shared by the buffered CSV scanners; parsing stays with each reader.</summary>
internal static class CsvSimdMasks
{
    // Saturating signed packing maps non-byte UTF-16 to zero or 255. Callers
    // admit delimiters 1..254, so Unicode cannot alias CSV punctuation.
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static Vector128<byte> PackText16Arm(ref ushort text, int position) =>
        AdvSimd.ExtractNarrowingSaturateUnsignedUpper(
            AdvSimd.ExtractNarrowingSaturateUnsignedLower(
                Vector128.LoadUnsafe(ref text, (nuint)position).AsInt16()),
            Vector128.LoadUnsafe(ref text, (nuint)(position + 8)).AsInt16());

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static uint ReadTextSpecial32Arm(ref ushort text, int position, byte delimiter, bool includeQuote)
    {
        var lower = Special(PackText16Arm(ref text, position), delimiter, includeQuote);
        var upper = Special(PackText16Arm(ref text, position + 16), delimiter, includeQuote);
        return Mask(lower, upper);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static Vector128<byte> Special(Vector128<byte> bytes, byte delimiter, bool includeQuote)
    {
        var matches = Vector128.Equals(bytes, Vector128.Create(delimiter))
            | Vector128.Equals(bytes, Vector128.Create((byte)'\r'))
            | Vector128.Equals(bytes, Vector128.Create((byte)'\n'));
        if (includeQuote) matches |= Vector128.Equals(bytes, Vector128.Create((byte)'"'));
        return matches;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void ReadText32Arm(ref ushort text, int position, byte delimiter,
        out uint delimiters, out uint quotes, out uint newlines) =>
        Read32(PackText16Arm(ref text, position), PackText16Arm(ref text, position + 16),
            delimiter, out delimiters, out quotes, out newlines);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void ReadBytes32(ref byte bytes, int position, byte delimiter,
        out uint delimiters, out uint quotes, out uint newlines) =>
        Read32(Vector128.LoadUnsafe(ref bytes, (nuint)position),
            Vector128.LoadUnsafe(ref bytes, (nuint)(position + 16)),
            delimiter, out delimiters, out quotes, out newlines);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static void Read32(Vector128<byte> lower, Vector128<byte> upper, byte delimiter,
        out uint delimiters, out uint quotes, out uint newlines)
    {
        var separator = Vector128.Create(delimiter);
        var quote = Vector128.Create((byte)'"');
        var cr = Vector128.Create((byte)'\r');
        var lf = Vector128.Create((byte)'\n');
        delimiters = Mask(Vector128.Equals(lower, separator), Vector128.Equals(upper, separator));
        quotes = Mask(Vector128.Equals(lower, quote), Vector128.Equals(upper, quote));
        newlines = Mask(Vector128.Equals(lower, cr) | Vector128.Equals(lower, lf),
            Vector128.Equals(upper, cr) | Vector128.Equals(upper, lf));
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static uint Mask(Vector128<byte> lower, Vector128<byte> upper) =>
        lower.ExtractMostSignificantBits() | (upper.ExtractMostSignificantBits() << 16);
}
#endif
