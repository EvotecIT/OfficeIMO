#if NET8_0_OR_GREATER
using System.Runtime.CompilerServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegReader {
    private static void InverseDctVector(int[] input, byte[] output, int[] workspace) {
        var r0 = LoadTransformRow(input, 0);
        var r1 = LoadTransformRow(input, 8);
        var r2 = LoadTransformRow(input, 16);
        var r3 = LoadTransformRow(input, 24);
        var r4 = LoadTransformRow(input, 32);
        var r5 = LoadTransformRow(input, 40);
        var r6 = LoadTransformRow(input, 48);
        var r7 = LoadTransformRow(input, 56);
        if (!CanTransformInt32(r0, r1, r2, r3, r4, r5, r6, r7)) {
            InverseDctScalar(input, output, workspace);
            return;
        }
        TransformIntegerVector(ref r0, ref r1, ref r2, ref r3, ref r4, ref r5, ref r6, ref r7, finalPass: false);
        var t0 = r0.AsSingle();
        var t1 = r1.AsSingle();
        var t2 = r2.AsSingle();
        var t3 = r3.AsSingle();
        var t4 = r4.AsSingle();
        var t5 = r5.AsSingle();
        var t6 = r6.AsSingle();
        var t7 = r7.AsSingle();
        OfficeJpegBlockTranspose.Apply(ref t0, ref t1, ref t2, ref t3, ref t4, ref t5, ref t6, ref t7);
        r0 = t0.AsInt32();
        r1 = t1.AsInt32();
        r2 = t2.AsInt32();
        r3 = t3.AsInt32();
        r4 = t4.AsInt32();
        r5 = t5.AsInt32();
        r6 = t6.AsInt32();
        r7 = t7.AsInt32();
        if (!CanTransformInt32(r0, r1, r2, r3, r4, r5, r6, r7)) {
            InverseDctScalar(input, output, workspace);
            return;
        }
        TransformIntegerVector(ref r0, ref r1, ref r2, ref r3, ref r4, ref r5, ref r6, ref r7, finalPass: true);
        t0 = ClampTransformPixels(r0); t1 = ClampTransformPixels(r1);
        t2 = ClampTransformPixels(r2); t3 = ClampTransformPixels(r3);
        t4 = ClampTransformPixels(r4); t5 = ClampTransformPixels(r5);
        t6 = ClampTransformPixels(r6); t7 = ClampTransformPixels(r7);
        OfficeJpegBlockTranspose.Apply(ref t0, ref t1, ref t2, ref t3, ref t4, ref t5, ref t6, ref t7);
        StorePixelRow(t0, output, 0); StorePixelRow(t1, output, 8);
        StorePixelRow(t2, output, 16); StorePixelRow(t3, output, 24);
        StorePixelRow(t4, output, 32); StorePixelRow(t5, output, 40);
        StorePixelRow(t6, output, 48); StorePixelRow(t7, output, 56);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static Vector256<int> LoadTransformRow(int[] input, int offset) =>
        Vector256.LoadUnsafe(ref input[offset]);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static void StorePixelRow(Vector256<float> row, byte[] output, int offset) {
        var pixels = row.AsInt32();
        var shorts = Vector128.Narrow(pixels.GetLower(), pixels.GetUpper()).AsUInt16();
        Vector128.Narrow(shorts, Vector128<ushort>.Zero).GetLower().StoreUnsafe(ref output[offset]);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static Vector256<float> ClampTransformPixels(Vector256<int> values) {
        var pixels = Avx2.Add(values, Vector256.Create(128));
        return Avx2.Min(Avx2.Max(pixels, Vector256<int>.Zero), Vector256.Create(255)).AsSingle();
    }

    // Every intermediate fits Int32 when |DC| <= 65535 and |AC| <= 8191:
    // the largest intermediate plus rounding is below one billion. Check both passes;
    // larger coefficients retain the existing 64-bit scalar transform.
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static bool CanTransformInt32(Vector256<int> r0, Vector256<int> r1,
        Vector256<int> r2, Vector256<int> r3, Vector256<int> r4,
        Vector256<int> r5, Vector256<int> r6, Vector256<int> r7) {
        var maximum = Avx2.Max(Avx2.Max(Avx2.Max(r1, r2), Avx2.Max(r3, r4)), Avx2.Max(Avx2.Max(r5, r6), r7));
        var minimum = Avx2.Min(Avx2.Min(Avx2.Min(r1, r2), Avx2.Min(r3, r4)), Avx2.Min(Avx2.Min(r5, r6), r7));
        var outside = Avx2.CompareGreaterThan(maximum, Vector256.Create(8191)) |
            Avx2.CompareGreaterThan(Vector256.Create(-8191), minimum) |
            Avx2.CompareGreaterThan(r0, Vector256.Create(65535)) |
            Avx2.CompareGreaterThan(Vector256.Create(-65535), r0);
        return Avx2.MoveMask(outside.AsByte()) == 0;
    }

    // Each lane is one column (or one row after transposition).
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static void TransformIntegerVector(ref Vector256<int> r0, ref Vector256<int> r1,
        ref Vector256<int> r2, ref Vector256<int> r3, ref Vector256<int> r4,
        ref Vector256<int> r5, ref Vector256<int> r6, ref Vector256<int> r7, bool finalPass) {
        var dcOnly = Vector256.Equals(r1 | r2 | r3 | r4 | r5 | r6 | r7, Vector256<int>.Zero);
        var dc = finalPass ? DescaleVector(r0, Pass1Bits + 3) : (r0 << Pass1Bits);
        if (Vector256.EqualsAll(dcOnly, Vector256<int>.AllBitsSet)) {
            r0 = r1 = r2 = r3 = r4 = r5 = r6 = r7 = dc;
            return;
        }
        var z1 = (r2 + r6) * Vector256.Create((int)Fix0_541196100);
        var tmp2 = z1 - r6 * Vector256.Create((int)Fix1_847759065);
        var tmp3 = z1 + r2 * Vector256.Create((int)Fix0_765366865);
        var tmp0 = (r0 + r4) << ConstBits;
        var tmp1 = (r0 - r4) << ConstBits;
        var tmp10 = tmp0 + tmp3;
        var tmp13 = tmp0 - tmp3;
        var tmp11 = tmp1 + tmp2;
        var tmp12 = tmp1 - tmp2;
        tmp0 = r7; tmp1 = r5; tmp2 = r3; tmp3 = r1;
        z1 = (tmp0 + tmp3) * Vector256.Create((int)-Fix0_899976223);
        var z2 = (tmp1 + tmp2) * Vector256.Create((int)-Fix2_562915447);
        var z3 = tmp0 + tmp2;
        var z4 = tmp1 + tmp3;
        var z5 = (z3 + z4) * Vector256.Create((int)Fix1_175875602);
        z3 = z3 * Vector256.Create((int)-Fix1_961570560) + z5;
        z4 = z4 * Vector256.Create((int)-Fix0_390180644) + z5;
        tmp0 = tmp0 * Vector256.Create((int)Fix0_298631336) + z1 + z3;
        tmp1 = tmp1 * Vector256.Create((int)Fix2_053119869) + z2 + z4;
        tmp2 = tmp2 * Vector256.Create((int)Fix3_072711026) + z2 + z3;
        tmp3 = tmp3 * Vector256.Create((int)Fix1_501321110) + z1 + z4;
        int shift = finalPass ? ConstBits + Pass1Bits + 3 : ConstBits - Pass1Bits;
        r0 = Vector256.ConditionalSelect(dcOnly, dc, DescaleVector(tmp10 + tmp3, shift));
        r7 = Vector256.ConditionalSelect(dcOnly, dc, DescaleVector(tmp10 - tmp3, shift));
        r1 = Vector256.ConditionalSelect(dcOnly, dc, DescaleVector(tmp11 + tmp2, shift));
        r6 = Vector256.ConditionalSelect(dcOnly, dc, DescaleVector(tmp11 - tmp2, shift));
        r2 = Vector256.ConditionalSelect(dcOnly, dc, DescaleVector(tmp12 + tmp1, shift));
        r5 = Vector256.ConditionalSelect(dcOnly, dc, DescaleVector(tmp12 - tmp1, shift));
        r3 = Vector256.ConditionalSelect(dcOnly, dc, DescaleVector(tmp13 + tmp0, shift));
        r4 = Vector256.ConditionalSelect(dcOnly, dc, DescaleVector(tmp13 - tmp0, shift));
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static Vector256<int> DescaleVector(Vector256<int> values, int shift) {
        var sign = values >> 31;
        var magnitude = (values ^ sign) - sign;
        var rounded = (magnitude + Vector256.Create(1 << (shift - 1))) >> shift;
        return (rounded ^ sign) - sign;
    }
}
#endif
