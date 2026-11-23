#if NET8_0_OR_GREATER
using System.Runtime.CompilerServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.Arm;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegReader {
    // AdvSimd form of InverseDctVector: the same Int32 transform on two four-column
    // halves, with the same range guards and the same scalar fallback, so every block
    // produces the samples InverseDctScalar produces.
    private static void InverseDctArm(int[] input, byte[] output, int[] workspace) {
        ref int source = ref input[0];
        var a0 = Vector128.LoadUnsafe(ref source, (nuint)0); var b0 = Vector128.LoadUnsafe(ref source, (nuint)4);
        var a1 = Vector128.LoadUnsafe(ref source, (nuint)8); var b1 = Vector128.LoadUnsafe(ref source, (nuint)12);
        var a2 = Vector128.LoadUnsafe(ref source, (nuint)16); var b2 = Vector128.LoadUnsafe(ref source, (nuint)20);
        var a3 = Vector128.LoadUnsafe(ref source, (nuint)24); var b3 = Vector128.LoadUnsafe(ref source, (nuint)28);
        var a4 = Vector128.LoadUnsafe(ref source, (nuint)32); var b4 = Vector128.LoadUnsafe(ref source, (nuint)36);
        var a5 = Vector128.LoadUnsafe(ref source, (nuint)40); var b5 = Vector128.LoadUnsafe(ref source, (nuint)44);
        var a6 = Vector128.LoadUnsafe(ref source, (nuint)48); var b6 = Vector128.LoadUnsafe(ref source, (nuint)52);
        var a7 = Vector128.LoadUnsafe(ref source, (nuint)56); var b7 = Vector128.LoadUnsafe(ref source, (nuint)60);
        if (!CanTransformInt32(a0, a1, a2, a3, a4, a5, a6, a7) || !CanTransformInt32(b0, b1, b2, b3, b4, b5, b6, b7)) {
            InverseDctScalar(input, output, workspace);
            return;
        }
        TransformIntegerVector(ref a0, ref a1, ref a2, ref a3, ref a4, ref a5, ref a6, ref a7, finalPass: false);
        TransformIntegerVector(ref b0, ref b1, ref b2, ref b3, ref b4, ref b5, ref b6, ref b7, finalPass: false);
        Transpose(ref a0, ref a1, ref a2, ref a3, ref a4, ref a5, ref a6, ref a7,
            ref b0, ref b1, ref b2, ref b3, ref b4, ref b5, ref b6, ref b7);
        if (!CanTransformInt32(a0, a1, a2, a3, a4, a5, a6, a7) || !CanTransformInt32(b0, b1, b2, b3, b4, b5, b6, b7)) {
            InverseDctScalar(input, output, workspace);
            return;
        }
        TransformIntegerVector(ref a0, ref a1, ref a2, ref a3, ref a4, ref a5, ref a6, ref a7, finalPass: true);
        TransformIntegerVector(ref b0, ref b1, ref b2, ref b3, ref b4, ref b5, ref b6, ref b7, finalPass: true);
        // Lanes hold rows after the first transpose; transpose back so each
        // vector pair is one output row.
        Transpose(ref a0, ref a1, ref a2, ref a3, ref a4, ref a5, ref a6, ref a7,
            ref b0, ref b1, ref b2, ref b3, ref b4, ref b5, ref b6, ref b7);
        ref byte target = ref output[0];
        StorePixelRow(a0, b0, ref target, 0); StorePixelRow(a1, b1, ref target, 8);
        StorePixelRow(a2, b2, ref target, 16); StorePixelRow(a3, b3, ref target, 24);
        StorePixelRow(a4, b4, ref target, 32); StorePixelRow(a5, b5, ref target, 40);
        StorePixelRow(a6, b6, ref target, 48); StorePixelRow(a7, b7, ref target, 56);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static void StorePixelRow(Vector128<int> low, Vector128<int> high, ref byte target, int offset) {
        var lowPixels = Vector128.Min(Vector128.Max(low + Vector128.Create(128), Vector128<int>.Zero), Vector128.Create(255));
        var highPixels = Vector128.Min(Vector128.Max(high + Vector128.Create(128), Vector128<int>.Zero), Vector128.Create(255));
        var shorts = Vector128.Narrow(lowPixels, highPixels).AsUInt16();
        var bytes = Vector128.Narrow(shorts, shorts);
        Unsafe.WriteUnaligned(ref Unsafe.Add(ref target, offset), bytes.AsUInt64().ToScalar());
    }

    // 8x8 transpose of the block held as rows of (low, high) four-lane halves.
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static void Transpose(
        ref Vector128<int> a0, ref Vector128<int> a1, ref Vector128<int> a2, ref Vector128<int> a3,
        ref Vector128<int> a4, ref Vector128<int> a5, ref Vector128<int> a6, ref Vector128<int> a7,
        ref Vector128<int> b0, ref Vector128<int> b1, ref Vector128<int> b2, ref Vector128<int> b3,
        ref Vector128<int> b4, ref Vector128<int> b5, ref Vector128<int> b6, ref Vector128<int> b7) {
        Transpose4(ref a0, ref a1, ref a2, ref a3);
        Transpose4(ref b0, ref b1, ref b2, ref b3);
        Transpose4(ref a4, ref a5, ref a6, ref a7);
        Transpose4(ref b4, ref b5, ref b6, ref b7);
        // Swap the off-diagonal quadrants.
        (b0, a4) = (a4, b0);
        (b1, a5) = (a5, b1);
        (b2, a6) = (a6, b2);
        (b3, a7) = (a7, b3);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static void Transpose4(ref Vector128<int> r0, ref Vector128<int> r1, ref Vector128<int> r2, ref Vector128<int> r3) {
        var t0 = AdvSimd.Arm64.TransposeEven(r0, r1).AsInt64();
        var t1 = AdvSimd.Arm64.TransposeOdd(r0, r1).AsInt64();
        var t2 = AdvSimd.Arm64.TransposeEven(r2, r3).AsInt64();
        var t3 = AdvSimd.Arm64.TransposeOdd(r2, r3).AsInt64();
        r0 = AdvSimd.Arm64.TransposeEven(t0, t2).AsInt32();
        r1 = AdvSimd.Arm64.TransposeEven(t1, t3).AsInt32();
        r2 = AdvSimd.Arm64.TransposeOdd(t0, t2).AsInt32();
        r3 = AdvSimd.Arm64.TransposeOdd(t1, t3).AsInt32();
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static bool CanTransformInt32(Vector128<int> r0, Vector128<int> r1,
        Vector128<int> r2, Vector128<int> r3, Vector128<int> r4,
        Vector128<int> r5, Vector128<int> r6, Vector128<int> r7) {
        var maximum = Vector128.Max(Vector128.Max(Vector128.Max(r1, r2), Vector128.Max(r3, r4)), Vector128.Max(Vector128.Max(r5, r6), r7));
        var minimum = Vector128.Min(Vector128.Min(Vector128.Min(r1, r2), Vector128.Min(r3, r4)), Vector128.Min(Vector128.Min(r5, r6), r7));
        return !(Vector128.GreaterThanAny(maximum, Vector128.Create(8191)) |
            Vector128.LessThanAny(minimum, Vector128.Create(-8191)) |
            Vector128.GreaterThanAny(r0, Vector128.Create(65535)) |
            Vector128.LessThanAny(r0, Vector128.Create(-65535)));
    }

    // Lane-for-lane copy of the Vector256 TransformIntegerVector.
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static void TransformIntegerVector(ref Vector128<int> r0, ref Vector128<int> r1,
        ref Vector128<int> r2, ref Vector128<int> r3, ref Vector128<int> r4,
        ref Vector128<int> r5, ref Vector128<int> r6, ref Vector128<int> r7, bool finalPass) {
        var dcOnly = Vector128.Equals(r1 | r2 | r3 | r4 | r5 | r6 | r7, Vector128<int>.Zero);
        var dc = finalPass ? DescaleVector(r0, Pass1Bits + 3) : (r0 << Pass1Bits);
        if (Vector128.EqualsAll(dcOnly, Vector128<int>.AllBitsSet)) {
            r0 = r1 = r2 = r3 = r4 = r5 = r6 = r7 = dc;
            return;
        }
        var z1 = (r2 + r6) * Vector128.Create((int)Fix0_541196100);
        var tmp2 = z1 - r6 * Vector128.Create((int)Fix1_847759065);
        var tmp3 = z1 + r2 * Vector128.Create((int)Fix0_765366865);
        var tmp0 = (r0 + r4) << ConstBits;
        var tmp1 = (r0 - r4) << ConstBits;
        var tmp10 = tmp0 + tmp3;
        var tmp13 = tmp0 - tmp3;
        var tmp11 = tmp1 + tmp2;
        var tmp12 = tmp1 - tmp2;
        tmp0 = r7; tmp1 = r5; tmp2 = r3; tmp3 = r1;
        z1 = (tmp0 + tmp3) * Vector128.Create((int)-Fix0_899976223);
        var z2 = (tmp1 + tmp2) * Vector128.Create((int)-Fix2_562915447);
        var z3 = tmp0 + tmp2;
        var z4 = tmp1 + tmp3;
        var z5 = (z3 + z4) * Vector128.Create((int)Fix1_175875602);
        z3 = z3 * Vector128.Create((int)-Fix1_961570560) + z5;
        z4 = z4 * Vector128.Create((int)-Fix0_390180644) + z5;
        tmp0 = tmp0 * Vector128.Create((int)Fix0_298631336) + z1 + z3;
        tmp1 = tmp1 * Vector128.Create((int)Fix2_053119869) + z2 + z4;
        tmp2 = tmp2 * Vector128.Create((int)Fix3_072711026) + z2 + z3;
        tmp3 = tmp3 * Vector128.Create((int)Fix1_501321110) + z1 + z4;
        int shift = finalPass ? ConstBits + Pass1Bits + 3 : ConstBits - Pass1Bits;
        r0 = Vector128.ConditionalSelect(dcOnly, dc, DescaleVector(tmp10 + tmp3, shift));
        r7 = Vector128.ConditionalSelect(dcOnly, dc, DescaleVector(tmp10 - tmp3, shift));
        r1 = Vector128.ConditionalSelect(dcOnly, dc, DescaleVector(tmp11 + tmp2, shift));
        r6 = Vector128.ConditionalSelect(dcOnly, dc, DescaleVector(tmp11 - tmp2, shift));
        r2 = Vector128.ConditionalSelect(dcOnly, dc, DescaleVector(tmp12 + tmp1, shift));
        r5 = Vector128.ConditionalSelect(dcOnly, dc, DescaleVector(tmp12 - tmp1, shift));
        r3 = Vector128.ConditionalSelect(dcOnly, dc, DescaleVector(tmp13 + tmp0, shift));
        r4 = Vector128.ConditionalSelect(dcOnly, dc, DescaleVector(tmp13 - tmp0, shift));
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static Vector128<int> DescaleVector(Vector128<int> values, int shift) {
        var sign = values >> 31;
        var magnitude = (values ^ sign) - sign;
        var rounded = (magnitude + Vector128.Create(1 << (shift - 1))) >> shift;
        return (rounded ^ sign) - sign;
    }
}
#endif
