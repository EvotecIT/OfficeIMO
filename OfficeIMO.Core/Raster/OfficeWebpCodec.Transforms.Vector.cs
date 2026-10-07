#if NET8_0_OR_GREATER
using System.Runtime.CompilerServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;
using System.Threading;

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

    private static int RestoreVp8lGreenVector(uint[] pixels, CancellationToken cancellationToken) {
        Vector128<uint> channelMask = Vector128.Create(255U);
        Vector128<uint> redBlueMask = Vector128.Create(0x00FF00FFU);
        Vector128<uint> retainedMask = Vector128.Create(0xFF00FF00U);
        int pixel = 0;
        for (; pixel <= pixels.Length - 4; pixel += 4) {
            if ((pixel & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            Vector128<uint> color = Vector128.LoadUnsafe(ref pixels[pixel]);
            Vector128<uint> green = Sse2.And(Sse2.ShiftRightLogical(color, 8), channelMask);
            Vector128<uint> addends = Sse2.Or(green, Sse2.ShiftLeftLogical(green, 16));
            // The red and blue sums fit independent 16-bit lanes. Masking
            // restores byte wraparound while alpha and green stay unchanged.
            Vector128<uint> redBlue = Sse2.And(Sse2.Add(Sse2.And(color, redBlueMask), addends), redBlueMask);
            Sse2.Or(Sse2.And(color, retainedMask), redBlue).StoreUnsafe(ref pixels[pixel]);
        }
        return pixel;
    }
}
#endif
