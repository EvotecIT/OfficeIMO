#if NET8_0_OR_GREATER
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;

namespace OfficeIMO.Drawing;

public static partial class OfficePngReader {
    // Cross-platform form of the RGBA8 filters for hosts without SSSE3 (for example
    // Arm64). Each pixel is one 32-bit word; byte lanes never carry into each other.
    private static void UnfilterRgba8Portable(Span<byte> current, ReadOnlySpan<byte> previous, int filter,
        CancellationToken cancellationToken) {
        switch (filter) {
            case 0:
                return;
            case 1:
                UnfilterRgbaSubPortable(current, cancellationToken);
                return;
            case 2:
                UnfilterRgbaUpPortable(current, previous, cancellationToken);
                return;
            case 3:
                UnfilterRgbaAveragePortable(current, previous, cancellationToken);
                return;
            case 4:
                UnfilterRgbaPaethPortable(current, previous, cancellationToken);
                return;
            default:
                throw new InvalidDataException("Unsupported PNG filter.");
        }
    }

    // Adds four independent bytes modulo 256.
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static uint AddBytes(uint left, uint right) =>
        ((left & 0x7F7F7F7FU) + (right & 0x7F7F7F7FU)) ^ ((left ^ right) & 0x80808080U);

    private static void UnfilterRgbaSubPortable(Span<byte> current, CancellationToken cancellationToken) {
        Span<uint> pixels = MemoryMarshal.Cast<byte, uint>(current);
        uint left = 0;
        for (int index = 0; index < pixels.Length; index++) {
            if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            left = AddBytes(pixels[index], left);
            pixels[index] = left;
        }
    }

    private static void UnfilterRgbaUpPortable(Span<byte> current, ReadOnlySpan<byte> previous,
        CancellationToken cancellationToken) {
        ref byte data = ref MemoryMarshal.GetReference(current);
        ref byte above = ref MemoryMarshal.GetReference(previous);
        int index = 0;
        for (; index <= current.Length - 16; index += 16) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            var sum = Vector128.LoadUnsafe(ref data, (nuint)index) + Vector128.LoadUnsafe(ref above, (nuint)index);
            sum.StoreUnsafe(ref data, (nuint)index);
        }
        for (; index < current.Length; index++) {
            current[index] = unchecked((byte)(current[index] + previous[index]));
        }
    }

    private static void UnfilterRgbaAveragePortable(Span<byte> current, ReadOnlySpan<byte> previous,
        CancellationToken cancellationToken) {
        Span<uint> pixels = MemoryMarshal.Cast<byte, uint>(current);
        ReadOnlySpan<uint> abovePixels = MemoryMarshal.Cast<byte, uint>(previous);
        uint left = 0;
        for (int index = 0; index < pixels.Length; index++) {
            if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            uint above = abovePixels[index];
            // Per-byte floor((left + above) / 2) without carries between bytes.
            uint prediction = (left & above) + (((left ^ above) & 0xFEFEFEFEU) >> 1);
            left = AddBytes(pixels[index], prediction);
            pixels[index] = left;
        }
    }

    private static void UnfilterRgbaPaethPortable(Span<byte> current, ReadOnlySpan<byte> previous,
        CancellationToken cancellationToken) {
        Span<uint> pixels = MemoryMarshal.Cast<byte, uint>(current);
        ReadOnlySpan<uint> abovePixels = MemoryMarshal.Cast<byte, uint>(previous);
        var left = Vector128<short>.Zero;
        var upperLeft = Vector128<short>.Zero;
        for (int index = 0; index < pixels.Length; index++) {
            if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            var above = WidenPixel(abovePixels[index]);
            // Same selection as the SSE path: high below the low threshold, low above
            // the high threshold, the diagonal between them, with PNG's tie order.
            var low = Vector128.Min(left, above);
            var high = Vector128.Max(left, above);
            var threshold = upperLeft + upperLeft + upperLeft - (left + above);
            var aboveLow = Vector128.GreaterThan(threshold, low);
            var belowHigh = Vector128.GreaterThan(high, threshold);
            var prediction = Vector128.ConditionalSelect(belowHigh,
                Vector128.ConditionalSelect(aboveLow, upperLeft, high), low);
            uint predicted = Vector128.Narrow(prediction, prediction).AsUInt32().ToScalar();
            uint value = AddBytes(pixels[index], predicted);
            pixels[index] = value;
            left = WidenPixel(value);
            upperLeft = above;
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static Vector128<short> WidenPixel(uint pixel) =>
        Vector128.WidenLower(Vector128.CreateScalar(pixel).AsByte()).AsInt16();
}
#endif
