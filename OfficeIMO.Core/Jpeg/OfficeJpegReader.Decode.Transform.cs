using System;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegReader {
    private const int ConstBits = 13;
    private const int Pass1Bits = 2;
    // Fixed-point constants from the IJG islow integer IDCT implementation.
    private const long Fix0_298631336 = 2446;
    private const long Fix0_390180644 = 3196;
    private const long Fix0_541196100 = 4433;
    private const long Fix0_765366865 = 6270;
    private const long Fix0_899976223 = 7373;
    private const long Fix1_175875602 = 9633;
    private const long Fix1_501321110 = 12299;
    private const long Fix1_847759065 = 15137;
    private const long Fix1_961570560 = 16069;
    private const long Fix2_053119869 = 16819;
    private const long Fix2_562915447 = 20995;
    private const long Fix3_072711026 = 25172;

    private static void InverseDct(int[] input, byte[] output, int[] workspace) {
#if NET8_0_OR_GREATER
        if (System.Runtime.Intrinsics.X86.Avx2.IsSupported) {
            InverseDctVector(input, output, workspace);
            return;
        }
#endif
        InverseDctScalar(input, output, workspace);
    }

    private static void InverseDctScalar(int[] input, byte[] output, int[] workspace) {

        // Pass 1: process columns into the workspace (scaled by Pass1Bits).
        for (var ctr = 0; ctr < 8; ctr++) {
            var c0 = input[ctr];
            var c1 = input[ctr + 8];
            var c2 = input[ctr + 16];
            var c3 = input[ctr + 24];
            var c4 = input[ctr + 32];
            var c5 = input[ctr + 40];
            var c6 = input[ctr + 48];
            var c7 = input[ctr + 56];

            if (c1 == 0 && c2 == 0 && c3 == 0 && c4 == 0 && c5 == 0 && c6 == 0 && c7 == 0) {
                var dc = c0 << Pass1Bits;
                workspace[ctr] = dc;
                workspace[ctr + 8] = dc;
                workspace[ctr + 16] = dc;
                workspace[ctr + 24] = dc;
                workspace[ctr + 32] = dc;
                workspace[ctr + 40] = dc;
                workspace[ctr + 48] = dc;
                workspace[ctr + 56] = dc;
                continue;
            }

            long tmp0;
            long tmp1;
            long tmp2;
            long tmp3;
            long tmp10;
            long tmp11;
            long tmp12;
            long tmp13;
            long z1;
            long z2;
            long z3;
            long z4;
            long z5;

            // Even part.
            z2 = c2;
            z3 = c6;
            z1 = (z2 + z3) * Fix0_541196100;
            tmp2 = z1 + z3 * -Fix1_847759065;
            tmp3 = z1 + z2 * Fix0_765366865;

            tmp0 = (c0 + c4) << ConstBits;
            tmp1 = (c0 - c4) << ConstBits;

            tmp10 = tmp0 + tmp3;
            tmp13 = tmp0 - tmp3;
            tmp11 = tmp1 + tmp2;
            tmp12 = tmp1 - tmp2;

            // Odd part.
            tmp0 = c7;
            tmp1 = c5;
            tmp2 = c3;
            tmp3 = c1;

            z1 = tmp0 + tmp3;
            z2 = tmp1 + tmp2;
            z3 = tmp0 + tmp2;
            z4 = tmp1 + tmp3;
            z5 = (z3 + z4) * Fix1_175875602;

            tmp0 *= Fix0_298631336;
            tmp1 *= Fix2_053119869;
            tmp2 *= Fix3_072711026;
            tmp3 *= Fix1_501321110;
            z1 *= -Fix0_899976223;
            z2 *= -Fix2_562915447;
            z3 *= -Fix1_961570560;
            z4 *= -Fix0_390180644;

            z3 += z5;
            z4 += z5;

            tmp0 += z1 + z3;
            tmp1 += z2 + z4;
            tmp2 += z2 + z3;
            tmp3 += z1 + z4;

            workspace[ctr] = Descale(tmp10 + tmp3, ConstBits - Pass1Bits);
            workspace[ctr + 56] = Descale(tmp10 - tmp3, ConstBits - Pass1Bits);
            workspace[ctr + 8] = Descale(tmp11 + tmp2, ConstBits - Pass1Bits);
            workspace[ctr + 48] = Descale(tmp11 - tmp2, ConstBits - Pass1Bits);
            workspace[ctr + 16] = Descale(tmp12 + tmp1, ConstBits - Pass1Bits);
            workspace[ctr + 40] = Descale(tmp12 - tmp1, ConstBits - Pass1Bits);
            workspace[ctr + 24] = Descale(tmp13 + tmp0, ConstBits - Pass1Bits);
            workspace[ctr + 32] = Descale(tmp13 - tmp0, ConstBits - Pass1Bits);
        }

        // Pass 2: process rows from the workspace into final pixels.
        for (var ctr = 0; ctr < 8; ctr++) {
            var row = ctr * 8;
            var w0 = workspace[row];
            var w1 = workspace[row + 1];
            var w2 = workspace[row + 2];
            var w3 = workspace[row + 3];
            var w4 = workspace[row + 4];
            var w5 = workspace[row + 5];
            var w6 = workspace[row + 6];
            var w7 = workspace[row + 7];

            if (w1 == 0 && w2 == 0 && w3 == 0 && w4 == 0 && w5 == 0 && w6 == 0 && w7 == 0) {
                var dc = Descale(w0, Pass1Bits + 3) + 128;
                var clamped = ClampToByte(dc);
                output[row] = clamped;
                output[row + 1] = clamped;
                output[row + 2] = clamped;
                output[row + 3] = clamped;
                output[row + 4] = clamped;
                output[row + 5] = clamped;
                output[row + 6] = clamped;
                output[row + 7] = clamped;
                continue;
            }

            long tmp0;
            long tmp1;
            long tmp2;
            long tmp3;
            long tmp10;
            long tmp11;
            long tmp12;
            long tmp13;
            long z1;
            long z2;
            long z3;
            long z4;
            long z5;

            // Even part.
            z2 = w2;
            z3 = w6;
            z1 = (z2 + z3) * Fix0_541196100;
            tmp2 = z1 + z3 * -Fix1_847759065;
            tmp3 = z1 + z2 * Fix0_765366865;

            tmp0 = (w0 + w4) << ConstBits;
            tmp1 = (w0 - w4) << ConstBits;

            tmp10 = tmp0 + tmp3;
            tmp13 = tmp0 - tmp3;
            tmp11 = tmp1 + tmp2;
            tmp12 = tmp1 - tmp2;

            // Odd part.
            tmp0 = w7;
            tmp1 = w5;
            tmp2 = w3;
            tmp3 = w1;

            z1 = tmp0 + tmp3;
            z2 = tmp1 + tmp2;
            z3 = tmp0 + tmp2;
            z4 = tmp1 + tmp3;
            z5 = (z3 + z4) * Fix1_175875602;

            tmp0 *= Fix0_298631336;
            tmp1 *= Fix2_053119869;
            tmp2 *= Fix3_072711026;
            tmp3 *= Fix1_501321110;
            z1 *= -Fix0_899976223;
            z2 *= -Fix2_562915447;
            z3 *= -Fix1_961570560;
            z4 *= -Fix0_390180644;

            z3 += z5;
            z4 += z5;

            tmp0 += z1 + z3;
            tmp1 += z2 + z4;
            tmp2 += z2 + z3;
            tmp3 += z1 + z4;

            var shift = ConstBits + Pass1Bits + 3;
            output[row] = ClampToByte(Descale(tmp10 + tmp3, shift) + 128);
            output[row + 7] = ClampToByte(Descale(tmp10 - tmp3, shift) + 128);
            output[row + 1] = ClampToByte(Descale(tmp11 + tmp2, shift) + 128);
            output[row + 6] = ClampToByte(Descale(tmp11 - tmp2, shift) + 128);
            output[row + 2] = ClampToByte(Descale(tmp12 + tmp1, shift) + 128);
            output[row + 5] = ClampToByte(Descale(tmp12 - tmp1, shift) + 128);
            output[row + 3] = ClampToByte(Descale(tmp13 + tmp0, shift) + 128);
            output[row + 4] = ClampToByte(Descale(tmp13 - tmp0, shift) + 128);
        }
    }

    private static int Descale(long value, int shift) {
        if (shift <= 0) return (int)value;
        var round = 1L << (shift - 1);
        if (value >= 0) {
            return (int)((value + round) >> shift);
        }
        return (int)(-(((-value) + round) >> shift));
    }

}
