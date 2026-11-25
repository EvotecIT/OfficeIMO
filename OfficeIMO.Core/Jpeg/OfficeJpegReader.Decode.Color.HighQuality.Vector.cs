#if NET8_0_OR_GREATER
using System.Runtime.CompilerServices;
using System.Runtime.Intrinsics;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegReader {
    // Vector form of ComposeYccPixels for RGBA rows whose chroma columns are full or
    // half resolution. Sixteen pixels per step: chroma is interpolated in sixteenths with
    // the same integer weights, pixels whose chroma is whole use the integer tables'
    // formulas, and the rest use 22-bit fixed-point terms (factor * 2^18 per sixteenth,
    // split into an integer part and a 12-bit remainder). A fixed-point sum within a few
    // units of a rounding tie is redone by ComposeYccPixels, so every pixel matches the
    // double-precision path; the terms were checked against it for every luma and chroma
    // value.
    private const int YccFractionBits = 22;
    private const int YccTieMargin = 4;
    private const int RedTermHigh = 367525, RedTermLow = 3637;
    private const int BlueTermHigh = 464519, BlueTermLow = 688;
    private const int GreenCbTermHigh = 90213, GreenCbTermLow = 1075;
    private const int GreenCrTermHigh = 187206, GreenCrTermLow = 2222;

    private static void ComposeYccRowVector(byte[] output, int target, int width, byte[] luma, int lumaRow,
        in ChromaRows cb, in ChromaRows cr, out int start, out int end) {
        start = 0;
        end = 0;
        // Packed pixels are stored as little-endian Int32 lanes.
        if (!Vector128.IsHardwareAccelerated || !BitConverter.IsLittleEndian) return;
        if (cb.Columns.Step is not (1 or 2) || cr.Columns.Step is not (1 or 2)) return;
        // Half-resolution chunks start on an even column with a chroma sample to their left.
        int first = cb.Columns.Step == 2 || cr.Columns.Step == 2 ? 2 : 0;
        int x = first;
        for (; x + 16 <= width; x += 16) {
            if (!CanLoadChroma(cb, x) || !CanLoadChroma(cr, x) ||
                lumaRow + x + 16 > luma.Length || target + (x + 16) * 4 > output.Length) {
                break;
            }
            LoadChroma16(cb, x, out Vector128<short> cbLow, out Vector128<short> cbHigh);
            LoadChroma16(cr, x, out Vector128<short> crLow, out Vector128<short> crHigh);
            Vector128<byte> lumaBytes = Vector128.LoadUnsafe(ref luma[0], (nuint)(lumaRow + x));
            ref byte destination = ref output[target + x * 4];
            int fallback = ConvertYcc8(Vector128.WidenLower(lumaBytes).AsInt16(), cbLow, crLow, ref destination) |
                (ConvertYcc8(Vector128.WidenUpper(lumaBytes).AsInt16(), cbHigh, crHigh, ref Unsafe.Add(ref destination, 32)) << 8);
            while (fallback != 0) {
                int lane = System.Numerics.BitOperations.TrailingZeroCount(fallback);
                fallback &= fallback - 1;
                ComposeYccPixels(output, target, outputRgba: true, luma, lumaRow, cb, cr, x + lane, x + lane + 1);
            }
        }
        if (x == first) return;
        start = first;
        end = x;
    }

    private static bool CanLoadChroma(in ChromaRows rows, int x) {
        if (rows.Columns.Step == 1) {
            return x + 16 <= rows.Columns.ComponentLength &&
                rows.Top + x + 16 <= rows.TopSamples.Length && rows.Bottom + x + 16 <= rows.BottomSamples.Length;
        }
        // Samples m - 1 through m + 8 are interpolated; sixteen bytes are loaded from m - 1.
        int sample = x >> 1;
        return sample >= 1 && sample + 8 < rows.Columns.ComponentLength &&
            rows.Top + sample + 15 <= rows.TopSamples.Length && rows.Bottom + sample + 15 <= rows.BottomSamples.Length;
    }

    // Chroma in sixteenths for output columns x..x+15, as two eight-lane halves.
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static void LoadChroma16(in ChromaRows rows, int x, out Vector128<short> low, out Vector128<short> high) {
        InterpolateColumns(rows, rows.TopSamples, rows.Top, x, out Vector128<short> topLow, out Vector128<short> topHigh);
        InterpolateColumns(rows, rows.BottomSamples, rows.Bottom, x, out Vector128<short> bottomLow, out Vector128<short> bottomHigh);
        Vector128<short> weight = Vector128.Create((short)rows.RowWeight);
        low = (topLow << 2) + (bottomLow - topLow) * weight;
        high = (topHigh << 2) + (bottomHigh - topHigh) * weight;
    }

    // Column interpolation in quarters: 4 * s for full resolution, 3 * s[m] + s[m -/+ 1]
    // (the quarter weights of the bilinear axis) for half resolution.
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static void InterpolateColumns(in ChromaRows rows, byte[] samples, int row, int x,
        out Vector128<short> low, out Vector128<short> high) {
        if (rows.Columns.Step == 1) {
            Vector128<byte> values = Vector128.LoadUnsafe(ref samples[0], (nuint)(row + x));
            low = Vector128.WidenLower(values).AsInt16() << 2;
            high = Vector128.WidenUpper(values).AsInt16() << 2;
            return;
        }
        Vector128<byte> window = Vector128.LoadUnsafe(ref samples[0], (nuint)(row + (x >> 1) - 1));
        Vector128<byte> nearest = Vector128.Shuffle(window,
            Vector128.Create((byte)1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8));
        Vector128<byte> neighbor = Vector128.Shuffle(window,
            Vector128.Create((byte)0, 2, 1, 3, 2, 4, 3, 5, 4, 6, 5, 7, 6, 8, 7, 9));
        low = Vector128.WidenLower(nearest).AsInt16() * Vector128.Create((short)3) + Vector128.WidenLower(neighbor).AsInt16();
        high = Vector128.WidenUpper(nearest).AsInt16() * Vector128.Create((short)3) + Vector128.WidenUpper(neighbor).AsInt16();
    }

    // Converts eight pixels to RGBA; returns the lanes that need the double expression.
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static int ConvertYcc8(Vector128<short> luma, Vector128<short> cb16, Vector128<short> cr16, ref byte destination) =>
        ConvertYcc4(Vector128.WidenLower(luma), Vector128.WidenLower(cb16), Vector128.WidenLower(cr16), ref destination) |
        (ConvertYcc4(Vector128.WidenUpper(luma), Vector128.WidenUpper(cb16), Vector128.WidenUpper(cr16), ref Unsafe.Add(ref destination, 16)) << 4);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static int ConvertYcc4(Vector128<int> luma, Vector128<int> cb16, Vector128<int> cr16, ref byte destination) {
        Vector128<int> whole = Vector128.Equals((cb16 | cr16) & Vector128.Create(15), Vector128<int>.Zero);

        // Whole chroma: CrToR, CbToG + CrToG and CbToB.
        Vector128<int> cbCentered = (cb16 >> 4) - Vector128.Create(128);
        Vector128<int> crCentered = (cr16 >> 4) - Vector128.Create(128);
        Vector128<int> wholeRed = luma + ((crCentered * Vector128.Create(91881) + Vector128.Create(32768)) >> 16);
        Vector128<int> wholeGreen = luma + ((cbCentered * Vector128.Create(-22554) + Vector128.Create(32768) +
            crCentered * Vector128.Create(-46802)) >> 16);
        Vector128<int> wholeBlue = luma + ((cbCentered * Vector128.Create(116130) + Vector128.Create(32768)) >> 16);

        // Fractional chroma: fixed-point luma + term + one half.
        Vector128<int> cbDelta = cb16 - Vector128.Create(2048);
        Vector128<int> crDelta = cr16 - Vector128.Create(2048);
        Vector128<int> scaledLuma = (luma << YccFractionBits) + Vector128.Create(1 << (YccFractionBits - 1));
        Vector128<int> red = scaledLuma + Term(crDelta, RedTermHigh, RedTermLow);
        Vector128<int> green = scaledLuma - Term(cbDelta, GreenCbTermHigh, GreenCbTermLow) -
            Term(crDelta, GreenCrTermHigh, GreenCrTermLow);
        Vector128<int> blue = scaledLuma + Term(cbDelta, BlueTermHigh, BlueTermLow);
        Vector128<int> tie = NearTie(red) | NearTie(green) | NearTie(blue);

        Vector128<int> redOut = Vector128.ConditionalSelect(whole, wholeRed, red >> YccFractionBits);
        Vector128<int> greenOut = Vector128.ConditionalSelect(whole, wholeGreen, green >> YccFractionBits);
        Vector128<int> blueOut = Vector128.ConditionalSelect(whole, wholeBlue, blue >> YccFractionBits);
        Vector128<int> pixels = ClampSample(redOut) | (ClampSample(greenOut) << 8) | (ClampSample(blueOut) << 16) |
            Vector128.Create(unchecked((int)0xFF000000));
        pixels.AsByte().StoreUnsafe(ref destination);

        Vector128<int> fallback = Vector128.AndNot(tie, whole);
        return fallback == Vector128<int>.Zero ? 0 : (int)Vector128.ExtractMostSignificantBits(fallback);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static Vector128<int> Term(Vector128<int> delta, int high, int low) =>
        delta * Vector128.Create(high) + ((delta * Vector128.Create(low)) >> 12);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static Vector128<int> NearTie(Vector128<int> value) {
        const int fractionMask = (1 << YccFractionBits) - 1;
        Vector128<uint> fraction = ((value & Vector128.Create(fractionMask)) - Vector128.Create(YccTieMargin)).AsUInt32();
        return Vector128.GreaterThan(fraction, Vector128.Create((uint)(fractionMask - 2 * YccTieMargin))).AsInt32();
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static Vector128<int> ClampSample(Vector128<int> value) =>
        Vector128.Min(Vector128.Max(value, Vector128<int>.Zero), Vector128.Create(255));
}
#endif
