using System;

namespace OfficeIMO.Drawing;

/// <summary>Orthonormal 8 × 8 JPEG DCT with a caller-owned, per-encode workspace.</summary>
internal static partial class OfficeJpegForwardTransform {
    internal static void Quantize(int[] input, int[] quantization, int[] output, double[] workspace) {
#if NET8_0_OR_GREATER
        if (System.Runtime.Intrinsics.X86.Avx2.IsSupported) {
            QuantizeVector(input, quantization, output);
            return;
        }
#endif
        QuantizeScalar(input, quantization, output, workspace);
    }

    internal static void QuantizeScalar(int[] input, int[] quantization, int[] output, double[] workspace) {
        for (int index = 0; index < 64; index++) workspace[index] = input[index];
        for (int row = 0; row < 8; row++) TransformLine(workspace, row * 8, 1);
        for (int column = 0; column < 8; column++) TransformLine(workspace, column, 8);
        for (int index = 0; index < 64; index++) {
            output[index] = (int)Math.Round(workspace[index] / quantization[index]);
        }
    }

    // Pair symmetric samples before applying cosines. The even frequencies use
    // four sums, the odd frequencies four differences, and both passes include
    // the orthonormal 1/2 scaling (1/sqrt(8) for DC).
    private static void TransformLine(double[] values, int offset, int stride) {
        const double dc = 0.3535533905932737622;
        const double c1 = 0.4903926402016152246;
        const double c2 = 0.4619397662556433781;
        const double c3 = 0.4157348061512726185;
        const double c5 = 0.2777851165098011124;
        const double c6 = 0.1913417161825448859;
        const double c7 = 0.0975451610080641339;

        double a0 = values[offset] + values[offset + 7 * stride];
        double a1 = values[offset + stride] + values[offset + 6 * stride];
        double a2 = values[offset + 2 * stride] + values[offset + 5 * stride];
        double a3 = values[offset + 3 * stride] + values[offset + 4 * stride];
        double b0 = values[offset] - values[offset + 7 * stride];
        double b1 = values[offset + stride] - values[offset + 6 * stride];
        double b2 = values[offset + 2 * stride] - values[offset + 5 * stride];
        double b3 = values[offset + 3 * stride] - values[offset + 4 * stride];

        double even0 = a0 + a3, even1 = a1 + a2;
        double even2 = a0 - a3, even3 = a1 - a2;
        values[offset] = (even0 + even1) * dc;
        values[offset + 4 * stride] = (even0 - even1) * dc;
        values[offset + 2 * stride] = even2 * c2 + even3 * c6;
        values[offset + 6 * stride] = even2 * c6 - even3 * c2;
        values[offset + stride] = b0 * c1 + b1 * c3 + b2 * c5 + b3 * c7;
        values[offset + 3 * stride] = b0 * c3 - b1 * c7 - b2 * c1 - b3 * c5;
        values[offset + 5 * stride] = b0 * c5 - b1 * c1 + b2 * c7 + b3 * c3;
        values[offset + 7 * stride] = b0 * c7 - b1 * c5 + b2 * c3 - b3 * c1;
    }
}
