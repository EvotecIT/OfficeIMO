// Adapted from CodeGlyphX, commit fc25e2fcf795d9c9a09b88708c47bdfeed5c446d.
// Copyright CodeGlyphX contributors. Apache-2.0; see THIRD-PARTY-NOTICES.md.
// OfficeIMO adaptation removes diagnostic scaffolding and adds bounded cancellation/resource handling.
namespace OfficeIMO.Drawing;

/// <summary>Integer reconstruction transforms shared by VP8 encoding and decoding.</summary>
internal static class OfficeVp8Transform {
    private const int BlockSize = 4;
    private const int CoefficientsPerBlock = 16;
    private const int IdctCospi8Sqrt2Minus1 = 20091;
    private const int IdctSinpi8Sqrt2 = 35468;
    internal static int[] InverseTransform4x4(int[] input) {
        var output = new int[CoefficientsPerBlock];
        var temp = new int[CoefficientsPerBlock];
        InverseTransform4x4(input, temp, output);
        return output;
    }

    internal static void InverseTransform4x4(int[] input, int[] temp, int[] output) {

        for (var i = 0; i < BlockSize; i++) {
            var ip0 = input[i];
            var ip4 = input[i + 4];
            var ip8 = input[i + 8];
            var ip12 = input[i + 12];

            var a1 = ip0 + ip8;
            var b1 = ip0 - ip8;
            var temp1 = (ip4 * IdctSinpi8Sqrt2) >> 16;
            var temp2 = ip12 + ((ip12 * IdctCospi8Sqrt2Minus1) >> 16);
            var c1 = temp1 - temp2;
            temp1 = ip4 + ((ip4 * IdctCospi8Sqrt2Minus1) >> 16);
            temp2 = (ip12 * IdctSinpi8Sqrt2) >> 16;
            var d1 = temp1 + temp2;

            temp[i] = unchecked((short)(a1 + d1));
            temp[i + 12] = unchecked((short)(a1 - d1));
            temp[i + 4] = unchecked((short)(b1 + c1));
            temp[i + 8] = unchecked((short)(b1 - c1));
        }

        for (var i = 0; i < BlockSize; i++) {
            var baseIndex = i * BlockSize;
            var t0 = temp[baseIndex];
            var t1 = temp[baseIndex + 1];
            var t2 = temp[baseIndex + 2];
            var t3 = temp[baseIndex + 3];

            var a1 = t0 + t2;
            var b1 = t0 - t2;
            var temp1 = (t1 * IdctSinpi8Sqrt2) >> 16;
            var temp2 = t3 + ((t3 * IdctCospi8Sqrt2Minus1) >> 16);
            var c1 = temp1 - temp2;
            temp1 = t1 + ((t1 * IdctCospi8Sqrt2Minus1) >> 16);
            temp2 = (t3 * IdctSinpi8Sqrt2) >> 16;
            var d1 = temp1 + temp2;

            output[baseIndex] = unchecked((short)((a1 + d1 + 4) >> 3));
            output[baseIndex + 3] = unchecked((short)((a1 - d1 + 4) >> 3));
            output[baseIndex + 1] = unchecked((short)((b1 + c1 + 4) >> 3));
            output[baseIndex + 2] = unchecked((short)((b1 - c1 + 4) >> 3));
        }

    }

    internal static int[] InverseWalshTransform4x4(int[] input) {
        var temp = new int[CoefficientsPerBlock];
        var output = new int[CoefficientsPerBlock];
        InverseWalshTransform4x4(input, temp, output);
        return output;
    }

    internal static void InverseWalshTransform4x4(int[] input, int[] temp, int[] output) {

        for (var i = 0; i < BlockSize; i++) {
            var ip0 = input[i];
            var ip4 = input[i + 4];
            var ip8 = input[i + 8];
            var ip12 = input[i + 12];

            var a1 = ip0 + ip12;
            var b1 = ip4 + ip8;
            var c1 = ip4 - ip8;
            var d1 = ip0 - ip12;

            temp[i] = unchecked((short)(a1 + b1));
            temp[i + 4] = unchecked((short)(c1 + d1));
            temp[i + 8] = unchecked((short)(a1 - b1));
            temp[i + 12] = unchecked((short)(d1 - c1));
        }

        for (var i = 0; i < BlockSize; i++) {
            var baseIndex = i * BlockSize;
            var t0 = temp[baseIndex];
            var t1 = temp[baseIndex + 1];
            var t2 = temp[baseIndex + 2];
            var t3 = temp[baseIndex + 3];

            var a1 = t0 + t3;
            var b1 = t1 + t2;
            var c1 = t1 - t2;
            var d1 = t0 - t3;

            var a2 = a1 + b1;
            var b2 = c1 + d1;
            var c2 = a1 - b1;
            var d2 = d1 - c1;

            output[baseIndex] = unchecked((short)((a2 + 3) >> 3));
            output[baseIndex + 1] = unchecked((short)((b2 + 3) >> 3));
            output[baseIndex + 2] = unchecked((short)((c2 + 3) >> 3));
            output[baseIndex + 3] = unchecked((short)((d2 + 3) >> 3));
        }

    }
}
