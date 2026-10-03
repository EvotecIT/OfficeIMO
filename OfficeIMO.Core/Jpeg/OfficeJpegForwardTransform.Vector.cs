#if NET8_0_OR_GREATER
using System.Runtime.CompilerServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegForwardTransform {
    // Eight lanes transform all columns together. Transposition then exposes
    // the rows to the same butterfly, without per-block scratch allocation.
    private static void QuantizeVector(int[] input, int[] quantization, int[] output) {
        var r0 = Vector256.ConvertToSingle(Vector256.LoadUnsafe(ref input[0]));
        var r1 = Vector256.ConvertToSingle(Vector256.LoadUnsafe(ref input[8]));
        var r2 = Vector256.ConvertToSingle(Vector256.LoadUnsafe(ref input[16]));
        var r3 = Vector256.ConvertToSingle(Vector256.LoadUnsafe(ref input[24]));
        var r4 = Vector256.ConvertToSingle(Vector256.LoadUnsafe(ref input[32]));
        var r5 = Vector256.ConvertToSingle(Vector256.LoadUnsafe(ref input[40]));
        var r6 = Vector256.ConvertToSingle(Vector256.LoadUnsafe(ref input[48]));
        var r7 = Vector256.ConvertToSingle(Vector256.LoadUnsafe(ref input[56]));
        TransformVector(ref r0, ref r1, ref r2, ref r3, ref r4, ref r5, ref r6, ref r7);
        TransposeVector(ref r0, ref r1, ref r2, ref r3, ref r4, ref r5, ref r6, ref r7);
        TransformVector(ref r0, ref r1, ref r2, ref r3, ref r4, ref r5, ref r6, ref r7);
        TransposeVector(ref r0, ref r1, ref r2, ref r3, ref r4, ref r5, ref r6, ref r7);
        QuantizeVectorRow(r0, quantization, output, 0);
        QuantizeVectorRow(r1, quantization, output, 8);
        QuantizeVectorRow(r2, quantization, output, 16);
        QuantizeVectorRow(r3, quantization, output, 24);
        QuantizeVectorRow(r4, quantization, output, 32);
        QuantizeVectorRow(r5, quantization, output, 40);
        QuantizeVectorRow(r6, quantization, output, 48);
        QuantizeVectorRow(r7, quantization, output, 56);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static void QuantizeVectorRow(Vector256<float> values, int[] quantization, int[] output, int offset) {
        var quant = Vector256.ConvertToSingle(Vector256.LoadUnsafe(ref quantization[offset]));
        var rounded = Avx.RoundToNearestInteger(values / quant);
        Avx.ConvertToVector256Int32WithTruncation(rounded).StoreUnsafe(ref output[offset]);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static void TransformVector(ref Vector256<float> r0, ref Vector256<float> r1,
        ref Vector256<float> r2, ref Vector256<float> r3, ref Vector256<float> r4,
        ref Vector256<float> r5, ref Vector256<float> r6, ref Vector256<float> r7) {
        var dc = Vector256.Create(0.3535533905932737622f);
        var c1 = Vector256.Create(0.4903926402016152246f);
        var c2 = Vector256.Create(0.4619397662556433781f);
        var c3 = Vector256.Create(0.4157348061512726185f);
        var c5 = Vector256.Create(0.2777851165098011124f);
        var c6 = Vector256.Create(0.1913417161825448859f);
        var c7 = Vector256.Create(0.0975451610080641339f);
        var a0 = r0 + r7; var a1 = r1 + r6; var a2 = r2 + r5; var a3 = r3 + r4;
        var b0 = r0 - r7; var b1 = r1 - r6; var b2 = r2 - r5; var b3 = r3 - r4;
        var even0 = a0 + a3; var even1 = a1 + a2;
        var even2 = a0 - a3; var even3 = a1 - a2;
        r0 = (even0 + even1) * dc;
        r4 = (even0 - even1) * dc;
        r2 = even2 * c2 + even3 * c6;
        r6 = even2 * c6 - even3 * c2;
        r1 = b0 * c1 + b1 * c3 + b2 * c5 + b3 * c7;
        r3 = b0 * c3 - b1 * c7 - b2 * c1 - b3 * c5;
        r5 = b0 * c5 - b1 * c1 + b2 * c7 + b3 * c3;
        r7 = b0 * c7 - b1 * c5 + b2 * c3 - b3 * c1;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static void TransposeVector(ref Vector256<float> r0, ref Vector256<float> r1,
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
