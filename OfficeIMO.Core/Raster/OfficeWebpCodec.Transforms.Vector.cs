#if NET8_0_OR_GREATER
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;

namespace OfficeIMO.Drawing;

public static partial class OfficeWebpCodec {
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static uint SelectVp8lNeighborVector(uint left, uint top, uint topLeft) {
        Vector128<byte> corner = Vector128.CreateScalar(topLeft).AsByte();
        // |(left + top - corner) - left| equals |top - corner|;
        // the other distance is |left - corner|. The upper bytes are zero,
        // so SAD sums exactly the four independent ARGB channel distances.
        ulong leftDistance = Sse2.SumAbsoluteDifferences(Vector128.CreateScalar(top).AsByte(), corner).AsUInt64().GetElement(0);
        ulong topDistance = Sse2.SumAbsoluteDifferences(Vector128.CreateScalar(left).AsByte(), corner).AsUInt64().GetElement(0);
        return leftDistance < topDistance ? left : top;
    }

    /// <summary>Inverts the color transform four pixels at a time; returns the first unprocessed index.</summary>
    private static int ApplyVp8lColorTransformVector(Span<uint> pixels, int start, int end,
        int greenToRed, int greenToBlue, int redToBlue) {
        int index = start;
        if (!Vector128.IsHardwareAccelerated) return index;
        ref uint origin = ref MemoryMarshal.GetReference(pixels);
        Vector128<int> greenToRedVector = Vector128.Create(greenToRed);
        Vector128<int> greenToBlueVector = Vector128.Create(greenToBlue);
        Vector128<int> redToBlueVector = Vector128.Create(redToBlue);
        Vector128<int> channel = Vector128.Create(255);
        Vector128<int> retained = Vector128.Create(unchecked((int)0xFF00FF00U));
        for (; index <= end - 4; index += 4) {
            Vector128<int> color = Vector128.LoadUnsafe(ref origin, (nuint)index).AsInt32();
            // Shifting a channel to the top byte and back arithmetically sign-extends it.
            Vector128<int> green = Vector128.ShiftRightArithmetic(color << 16, 24);
            Vector128<int> red = (color >> 16) + Vector128.ShiftRightArithmetic(greenToRedVector * green, 5);
            Vector128<int> signedRed = Vector128.ShiftRightArithmetic(red << 24, 24);
            Vector128<int> blue = color + Vector128.ShiftRightArithmetic(greenToBlueVector * green, 5) +
                Vector128.ShiftRightArithmetic(redToBlueVector * signedRed, 5);
            ((color & retained) | ((red & channel) << 16) | (blue & channel)).AsUInt32().StoreUnsafe(ref origin, (nuint)index);
        }
        return index;
    }

    /// <summary>Adds green into red and blue for four pixels at a time; returns the first unprocessed index.</summary>
    private static int AddVp8lGreenVector(Span<uint> pixels, int start, int end) {
        int index = start;
        if (!Vector128.IsHardwareAccelerated) return index;
        Vector128<uint> channelMask = Vector128.Create(255U);
        Vector128<uint> redBlueMask = Vector128.Create(0x00FF00FFU);
        Vector128<uint> retainedMask = Vector128.Create(0xFF00FF00U);
        ref uint origin = ref MemoryMarshal.GetReference(pixels);
        for (; index <= end - 4; index += 4) {
            Vector128<uint> color = Vector128.LoadUnsafe(ref origin, (nuint)index);
            Vector128<uint> green = (color >> 8) & channelMask;
            // The red and blue sums fit independent 16-bit lanes. Masking
            // restores byte wraparound while alpha and green stay unchanged.
            Vector128<uint> redBlue = ((color & redBlueMask) + (green | (green << 16))) & redBlueMask;
            ((color & retainedMask) | redBlue).StoreUnsafe(ref origin, (nuint)index);
        }
        return index;
    }

    /// <summary>Writes ARGB words as RGBA bytes four pixels at a time; returns the first unprocessed index.</summary>
    private static int WriteVp8lRgbaVector(ReadOnlySpan<uint> argb, int start, int end, byte[] rgba) {
        int index = start;
        if (!Vector128.IsHardwareAccelerated) return index;
        // Little-endian ARGB words hold B, G, R, A bytes; swap B and R.
        var order = Vector128.Create((byte)2, 1, 0, 3, 6, 5, 4, 7, 10, 9, 8, 11, 14, 13, 12, 15);
        ref byte source = ref Unsafe.As<uint, byte>(ref MemoryMarshal.GetReference(argb));
        ref byte target = ref rgba[0];
        for (; index <= end - 4; index += 4) {
            Vector128<byte> bytes = Vector128.LoadUnsafe(ref source, (nuint)index * 4);
            Vector128.Shuffle(bytes, order).StoreUnsafe(ref target, (nuint)index * 4);
        }
        return index;
    }

    /// <summary>
    /// Restores four pixels at a time for predictor modes that read only the previous row, so no
    /// lane depends on another. Returns the first unprocessed index.
    /// </summary>
    private static int PredictVp8lTopRunVector(Span<uint> p, int mode, int start, int end, int width) {
        int index = start;
        if (!Vector128.IsHardwareAccelerated) return index;
        ref uint origin = ref MemoryMarshal.GetReference(p);
        Vector128<uint> halfMask = Vector128.Create(0xFEFEFEFEU);
        Vector128<uint> opaque = Vector128.Create(0xFF000000U);
        for (; index <= end - 4; index += 4) {
            nuint top = (nuint)(index - width);
            Vector128<uint> prediction;
            switch (mode) {
                case 0:
                    prediction = opaque;
                    break;
                case 2:
                    prediction = Vector128.LoadUnsafe(ref origin, top);
                    break;
                case 3:
                    prediction = Vector128.LoadUnsafe(ref origin, top + 1);
                    break;
                case 4:
                    prediction = Vector128.LoadUnsafe(ref origin, top - 1);
                    break;
                case 8:
                    prediction = Average(Vector128.LoadUnsafe(ref origin, top - 1), Vector128.LoadUnsafe(ref origin, top), halfMask);
                    break;
                default:
                    prediction = Average(Vector128.LoadUnsafe(ref origin, top), Vector128.LoadUnsafe(ref origin, top + 1), halfMask);
                    break;
            }
            Vector128<byte> residual = Vector128.LoadUnsafe(ref origin, (nuint)index).AsByte();
            (residual + prediction.AsByte()).AsUInt32().StoreUnsafe(ref origin, (nuint)index);
        }
        return index;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        static Vector128<uint> Average(Vector128<uint> first, Vector128<uint> second, Vector128<uint> halfMask) =>
            (first & second) + (((first ^ second) & halfMask) >> 1);
    }
}
#endif
