namespace OfficeIMO.Drawing;

public static partial class OfficeRasterResampler {
    /// <summary>Working pixels for <see cref="ReduceByIntegerBox"/>: its output plus one row of accumulators.</summary>
    internal static long GetIntegerBoxReductionPixelCost(int sourceWidth, int sourceHeight, int factorX, int factorY) {
        long width = (sourceWidth + (long)factorX - 1L) / factorX;
        long height = (sourceHeight + (long)factorY - 1L) / factorY;
        // Four double accumulators per output column occupy eight RGBA pixels.
        return checked(width * height + width * 8L);
    }

    /// <summary>
    /// Averages whole factorX by factorY source blocks in premultiplied encoded-sRGB space.
    /// Unlike the separable filters, it needs no full-height float scratch, so it can
    /// bring very large sources within a working budget before a final area resize.
    /// </summary>
    internal static OfficeRasterImage ReduceByIntegerBox(
        OfficeRasterImage source, int factorX, int factorY, CancellationToken cancellationToken) {
        if (factorX < 1) throw new ArgumentOutOfRangeException(nameof(factorX));
        if (factorY < 1) throw new ArgumentOutOfRangeException(nameof(factorY));
        int width = (int)((source.Width + (long)factorX - 1L) / factorX);
        int height = (int)((source.Height + (long)factorY - 1L) / factorY);
        byte[] input = source.PixelBuffer;
        byte[] output = new byte[checked(width * height * 4)];
        double[] sums = new double[checked(width * 4)];
        int sourceStride = source.Width * 4;
        for (int y = 0; y < height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            Array.Clear(sums, 0, sums.Length);
            int top = y * factorY;
            int bottom = Math.Min(source.Height, top + factorY);
            for (int sourceY = top; sourceY < bottom; sourceY++) {
                int row = sourceY * sourceStride;
                for (int sourceX = 0; sourceX < source.Width; sourceX++) {
                    int pixel = row + sourceX * 4;
                    double alpha = input[pixel + 3];
                    int sum = (sourceX / factorX) * 4;
                    sums[sum] += input[pixel] * alpha;
                    sums[sum + 1] += input[pixel + 1] * alpha;
                    sums[sum + 2] += input[pixel + 2] * alpha;
                    sums[sum + 3] += alpha;
                }
            }
            int rows = bottom - top;
            int target = y * width * 4;
            for (int x = 0; x < width; x++) {
                int sum = x * 4;
                double alphaSum = sums[sum + 3];
                if (alphaSum <= 0D) continue;
                int columns = Math.Min(source.Width, (x + 1) * factorX) - x * factorX;
                output[target + sum] = ToByte(sums[sum] / alphaSum);
                output[target + sum + 1] = ToByte(sums[sum + 1] / alphaSum);
                output[target + sum + 2] = ToByte(sums[sum + 2] / alphaSum);
                output[target + sum + 3] = ToByte(alphaSum / ((double)columns * rows));
            }
        }
        return OfficeRasterImage.FromOwnedRgba32(width, height, output);
    }
}
