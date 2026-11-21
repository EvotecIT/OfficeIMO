namespace OfficeIMO.Drawing;

public static partial class OfficeRasterFilters {
    /// <summary>
    /// Blurs one 8-bit coverage plane with the normalized Gaussian kernel of <see cref="GaussianBlur"/>, including
    /// its edge clamping. Soft shadows and glows need only their alpha, a quarter of the RGBA work.
    /// </summary>
    internal static byte[] GaussianBlurCoverage(byte[] coverage, int width, int height, double sigma, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(coverage);
        if ((long)width * height > coverage.LongLength) throw new ArgumentException("The coverage plane is shorter than its dimensions.", nameof(coverage));
        CalculateGaussianAdditionalWorkingBytes(width, height, sigma, false, out int radius);
        var weights = new float[radius * 2 + 1];
        double sum = 0D;
        for (int x = -radius; x <= radius; x++) sum += Math.Exp(-x * x / (2D * sigma * sigma));
        for (int x = -radius; x <= radius; x++) weights[x + radius] = (float)(Math.Exp(-x * x / (2D * sigma * sigma)) / sum);

        var horizontal = new float[width * height];
        for (int y = 0; y < height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            int row = y * width;
            for (int x = 0; x < width; x++) {
                float value = 0F;
                for (int k = 0; k < weights.Length; k++) {
                    int sx = x + k - radius;
                    sx = sx < 0 ? 0 : sx >= width ? width - 1 : sx;
                    value += coverage[row + sx] * weights[k];
                }
                horizontal[row + x] = value;
            }
        }

        var result = new byte[width * height];
        for (int y = 0; y < height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int x = 0; x < width; x++) {
                float value = 0F;
                for (int k = 0; k < weights.Length; k++) {
                    int sy = y + k - radius;
                    sy = sy < 0 ? 0 : sy >= height ? height - 1 : sy;
                    value += horizontal[sy * width + x] * weights[k];
                }
                result[y * width + x] = value >= 254.5F ? (byte)255 : value <= 0F ? (byte)0 : (byte)(value + 0.5F);
            }
        }
        return result;
    }
}
