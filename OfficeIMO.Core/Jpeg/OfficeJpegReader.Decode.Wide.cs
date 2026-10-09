using System;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegReader {
    // Wide quantization tables can exceed the fixed-point workspace range.
    // Keep ordinary blocks on the fast path and saturate only after the IDCT.
    private static void InverseDctWide(int[] input, byte[] output) {
        const double dcScale = 0.7071067811865475244;
        for (int y = 0; y < 8; y++) for (int x = 0; x < 8; x++) {
            double sum = 0;
            for (int v = 0; v < 8; v++) for (int u = 0; u < 8; u++)
                sum += input[v * 8 + u] * (u == 0 ? dcScale : 1) *
                    (v == 0 ? dcScale : 1) * IdctCos[x, u] * IdctCos[y, v];
            output[y * 8 + x] = (byte)Math.Max(0, Math.Min(255, Math.Round(sum / 4 + 128)));
        }
    }
}
