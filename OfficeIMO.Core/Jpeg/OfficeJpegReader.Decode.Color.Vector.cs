#if NET8_0_OR_GREATER
using System.Runtime.CompilerServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegReader {
    // Full-width luma and half-width chroma share each chroma value between two pixels.
    // Match scalar channel rounding: add both scaled green terms before rounding
    // the channel once and clamping. The caller handles any remaining pixels.
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    private static int ComposeYccToRgbaHalfChromaVector(
        byte[] luma, int lumaOffset, byte[] cb, int cbOffset, byte[] cr, int crOffset,
        byte[] output, int outputOffset, int width) {
        var duplicateChroma = Vector128.Create((byte)0, 0, 1, 1, 2, 2, 3, 3,
            128, 128, 128, 128, 128, 128, 128, 128);
        var center = Vector256.Create(128);
        var rounding = Vector256.Create(32768);
        var maximum = Vector256.Create(255);
        var alpha = Vector256.Create(unchecked((int)0xff000000));
        int x = 0;
        for (; x <= width - 8; x += 8) {
            // A complete group reads exactly eight luma and four samples per chroma
            // plane, then writes eight RGBA pixels. No padded input bytes are read.
            var yBytes = Vector128.CreateScalar(Unsafe.ReadUnaligned<ulong>(ref luma[lumaOffset + x])).AsByte();
            var cbBytes = Vector128.CreateScalar(Unsafe.ReadUnaligned<uint>(ref cb[cbOffset + x / 2])).AsByte();
            var crBytes = Vector128.CreateScalar(Unsafe.ReadUnaligned<uint>(ref cr[crOffset + x / 2])).AsByte();
            var y = Avx2.ConvertToVector256Int32(yBytes);
            var cbDelta = Avx2.Subtract(Avx2.ConvertToVector256Int32(Ssse3.Shuffle(cbBytes, duplicateChroma)), center);
            var crDelta = Avx2.Subtract(Avx2.ConvertToVector256Int32(Ssse3.Shuffle(crBytes, duplicateChroma)), center);
            var redContribution = Avx2.ShiftRightArithmetic(Avx2.Add(Avx2.MultiplyLow(crDelta, Vector256.Create(91881)), rounding), 16);
            var greenScaled = Avx2.Add(Avx2.MultiplyLow(cbDelta, Vector256.Create(-22554)), Avx2.MultiplyLow(crDelta, Vector256.Create(-46802)));
            var greenContribution = Avx2.ShiftRightArithmetic(Avx2.Add(greenScaled, rounding), 16);
            var blueContribution = Avx2.ShiftRightArithmetic(Avx2.Add(Avx2.MultiplyLow(cbDelta, Vector256.Create(116130)), rounding), 16);
            var red = Avx2.Min(Avx2.Max(Avx2.Add(y, redContribution), Vector256<int>.Zero), maximum);
            var green = Avx2.Min(Avx2.Max(Avx2.Add(y, greenContribution), Vector256<int>.Zero), maximum);
            var blue = Avx2.Min(Avx2.Max(Avx2.Add(y, blueContribution), Vector256<int>.Zero), maximum);
            var rgba = red | Avx2.ShiftLeftLogical(green, 8) | Avx2.ShiftLeftLogical(blue, 16) | alpha;
            rgba.AsByte().StoreUnsafe(ref output[outputOffset + x * 4]);
        }
        return x;
    }
}
#endif
