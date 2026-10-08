using System;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegReader {
    private static void WriteDctBlock(BaselineComponentState samples, int blockX, int blockY) {
        if (samples.WideBuffer == null) {
            InverseDct(samples.BlockCoeffs, samples.BlockPixels, samples.BlockWorkspace);
            WriteBlock(samples.Buffer, samples.Stride, blockX, blockY, samples.BlockPixels);
            return;
        }
        // Separable floating-point IDCT avoids overflowing the eight-bit fixed-point
        // workspace with twelve-bit coefficients or sixteen-bit quantization tables.
        // The guarded component workspace is reused for every block.
        const double dcScale = 0.7071067811865475244;
        double[] workspace = samples.WideWorkspace;
        int[] coefficients = samples.BlockCoeffs;
        for (int v = 0; v < 8; v++) for (int x = 0; x < 8; x++) {
            double sum = 0;
            for (int u = 0; u < 8; u++)
                sum += coefficients[v * 8 + u] * (u == 0 ? dcScale : 1) * IdctCos[x, u];
            workspace[v * 8 + x] = sum;
        }
        int midpoint = (samples.SampleMaximum + 1) / 2;
        for (int y = 0; y < 8; y++) for (int x = 0; x < 8; x++) {
            double sum = 0;
            for (int v = 0; v < 8; v++)
                sum += workspace[v * 8 + x] * (v == 0 ? dcScale : 1) * IdctCos[y, v];
            samples.WideBuffer[(blockY * 8 + y) * samples.Stride + blockX * 8 + x] =
                (ushort)Math.Max(0, Math.Min(samples.SampleMaximum, Math.Round(sum / 4 + midpoint)));
        }
    }
}
