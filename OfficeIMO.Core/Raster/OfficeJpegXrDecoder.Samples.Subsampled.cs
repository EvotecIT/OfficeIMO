using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegXrDecoder {
    // T.832 9.9.2: compact 2x2 / 2x4 chroma transforms.
    private static void InverseReducedTransform(int[] coefficients, int offset, int color, long[] work) {
        int count = color == 1 ? 4 : 8;
        for (int i = 0; i < count; i++) work[i] = coefficients[offset + i];
        if (color == 2) { work[0] -= (work[4] + 1) >> 1; work[4] += work[0]; }
        Hadamard(work, 0, 1, 2, 3, 0);
        long swap = work[1]; work[1] = work[2]; work[2] = swap;
        if (color == 2) {
            Hadamard(work, 4, 6, 5, 7, 0);
            swap = work[5]; work[5] = work[6]; work[6] = swap;
        }
        for (int i = 0; i < count; i++) coefficients[offset + i] = CheckedCoefficient(work[i]);
    }

    // T.832 9.10: vertical interpolation precedes horizontal interpolation.
    private static int[] UpsampleChroma(int[] samples, int width, int height, PlaneHeader plane, CancellationToken cancellation) {
        if (plane.Color == 1) {
            var vertical = new int[checked(width * height * 2)];
            for (int y = 0; y < height; y++) {
                cancellation.ThrowIfCancellationRequested();
                for (int x = 0; x < width; x++) {
                    if ((x & 4095) == 0) cancellation.ThrowIfCancellationRequested();
                    int current = samples[y * width + x];
                    vertical[y * 2 * width + x] = Interpolate(samples[Math.Max(0, y - 1) * width + x], current, plane.CenterY);
                    vertical[(y * 2 + 1) * width + x] = Interpolate(current, samples[Math.Min(height - 1, y + 1) * width + x], 4 + plane.CenterY);
                }
            }
            samples = vertical; height *= 2;
        }
        var output = new int[checked(width * height * 2)];
        for (int y = 0; y < height; y++) {
            cancellation.ThrowIfCancellationRequested();
            for (int x = 0; x < width; x++) {
                if ((x & 4095) == 0) cancellation.ThrowIfCancellationRequested();
                int current = samples[y * width + x];
                output[(y * width + x) * 2] = Interpolate(samples[y * width + Math.Max(0, x - 1)], current, plane.CenterX);
                output[(y * width + x) * 2 + 1] = Interpolate(current, samples[y * width + Math.Min(width - 1, x + 1)], 4 + plane.CenterX);
            }
        }
        return output;
    }

    private static int Interpolate(int first, int second, int firstWeight) =>
        CheckedCoefficient(((long)firstWeight * first + (8L - firstWeight) * second + 4) >> 3);
}
