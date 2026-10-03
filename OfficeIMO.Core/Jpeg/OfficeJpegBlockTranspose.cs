#if NET8_0_OR_GREATER
using System.Runtime.CompilerServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;

namespace OfficeIMO.Drawing;

// Shared bit-preserving 8x8 transpose for the JPEG forward and inverse transforms.
internal static class OfficeJpegBlockTranspose {
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void Apply(ref Vector256<float> r0, ref Vector256<float> r1,
        ref Vector256<float> r2, ref Vector256<float> r3, ref Vector256<float> r4,
        ref Vector256<float> r5, ref Vector256<float> r6, ref Vector256<float> r7) {
        var t0 = Avx.UnpackLow(r0, r1); var t1 = Avx.UnpackHigh(r0, r1);
        var t2 = Avx.UnpackLow(r2, r3); var t3 = Avx.UnpackHigh(r2, r3);
        var t4 = Avx.UnpackLow(r4, r5); var t5 = Avx.UnpackHigh(r4, r5);
        var t6 = Avx.UnpackLow(r6, r7); var t7 = Avx.UnpackHigh(r6, r7);
        var s0 = Avx.Shuffle(t0, t2, 0x44); var s1 = Avx.Shuffle(t0, t2, 0xEE);
        var s2 = Avx.Shuffle(t1, t3, 0x44); var s3 = Avx.Shuffle(t1, t3, 0xEE);
        var s4 = Avx.Shuffle(t4, t6, 0x44); var s5 = Avx.Shuffle(t4, t6, 0xEE);
        var s6 = Avx.Shuffle(t5, t7, 0x44); var s7 = Avx.Shuffle(t5, t7, 0xEE);
        r0 = Avx.Permute2x128(s0, s4, 0x20); r4 = Avx.Permute2x128(s0, s4, 0x31);
        r1 = Avx.Permute2x128(s1, s5, 0x20); r5 = Avx.Permute2x128(s1, s5, 0x31);
        r2 = Avx.Permute2x128(s2, s6, 0x20); r6 = Avx.Permute2x128(s2, s6, 0x31);
        r3 = Avx.Permute2x128(s3, s7, 0x20); r7 = Avx.Permute2x128(s3, s7, 0x31);
    }
}
#endif
