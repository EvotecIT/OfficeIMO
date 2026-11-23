namespace OfficeIMO.Drawing;

public static partial class OfficeRasterResampler {
    // Table-driven form of CopyBilinear for encoded sRGB. Source taps and fractions are
    // computed once per column and row with CopyBilinear's expressions, and each pixel
    // then evaluates the same weights, products and sums in the same order, so every
    // output byte equals the per-pixel path.
    private static void ResizeBilinearEncoded(
        byte[] input,
        int sourceWidth,
        int sourceHeight,
        byte[] output,
        int width,
        int height,
        double scaleX,
        double scaleY,
        CancellationToken cancellationToken) {
        var columns = BilinearAxis.Create(width, sourceWidth, scaleX, 4, cancellationToken);
        var rows = BilinearAxis.Create(height, sourceHeight, scaleY, sourceWidth * 4, cancellationToken);
        ResampleRows(height, (long)width * height * 4L, cancellationToken, y => {
            cancellationToken.ThrowIfCancellationRequested();
            int top = rows.First[y];
            int bottom = rows.Second[y];
            double ty = rows.Fraction[y];
            double inverseTy = rows.InverseFraction[y];
            int[] firstColumns = columns.First;
            int[] secondColumns = columns.Second;
            double[] fractions = columns.Fraction;
            double[] inverseFractions = columns.InverseFraction;
            int target = y * width * 4;
            for (int x = 0; x < width; x++, target += 4) {
                if ((x & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                double tx = fractions[x];
                double inverseTx = inverseFractions[x];
                double w00 = inverseTx * inverseTy;
                double w10 = tx * inverseTy;
                double w01 = inverseTx * ty;
                double w11 = tx * ty;
                int p00 = top + firstColumns[x];
                int p10 = top + secondColumns[x];
                int p01 = bottom + firstColumns[x];
                int p11 = bottom + secondColumns[x];
                int a00 = input[p00 + 3];
                int a10 = input[p10 + 3];
                int a01 = input[p01 + 3];
                int a11 = input[p11 + 3];
                double alpha = (a00 * w00) + (a10 * w10) + (a01 * w01) + (a11 * w11);
                if (alpha <= 0D) {
                    output[target] = output[target + 1] = output[target + 2] = output[target + 3] = 0;
                    continue;
                }
                for (int channel = 0; channel < 3; channel++) {
                    double value = ((double)(input[p00 + channel] * a00) * w00) +
                        ((double)(input[p10 + channel] * a10) * w10) +
                        ((double)(input[p01 + channel] * a01) * w01) +
                        ((double)(input[p11 + channel] * a11) * w11);
                    output[target + channel] = (byte)Math.Round(Clamp(value / alpha, 0D, 255D));
                }
                output[target + 3] = (byte)Math.Round(Clamp(alpha, 0D, 255D));
            }
        });
    }

    private sealed class BilinearAxis {
        public int[] First = Array.Empty<int>();
        public int[] Second = Array.Empty<int>();
        public double[] Fraction = Array.Empty<double>();
        public double[] InverseFraction = Array.Empty<double>();

        // Two int and two double entries per column and per row, plus array headers.
        public static long GetTableBytes(int width, int height) =>
            ((long)width + height) * (2L * sizeof(int) + 2L * sizeof(double)) + 8L * 24L;

        // Offsets are pre-multiplied by the axis stride in bytes.
        public static BilinearAxis Create(int length, int sourceLength, double scale, int stride,
            CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            var axis = new BilinearAxis {
                First = new int[length],
                Second = new int[length],
                Fraction = new double[length],
                InverseFraction = new double[length]
            };
            for (int i = 0; i < length; i++) {
                if ((i & 0xFFFF) == 0) cancellationToken.ThrowIfCancellationRequested();
                double position = ((i + 0.5D) * scale) - 0.5D;
                double sample = Clamp(position, 0D, sourceLength - 1D);
                int first = (int)Math.Floor(sample);
                int second = Clamp(first + 1, 0, sourceLength - 1);
                double fraction = sample - first;
                axis.First[i] = first * stride;
                axis.Second[i] = second * stride;
                axis.Fraction[i] = fraction;
                axis.InverseFraction[i] = 1D - fraction;
            }
            return axis;
        }
    }
}
