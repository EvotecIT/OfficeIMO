using System;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeRasterFilters {
    /// <summary>Applies a normalized Gaussian kernel in premultiplied RGBA space; sigma is measured in pixels.</summary>
    public static OfficeRasterImage GaussianBlur(OfficeRasterImage source, double sigma = 3D, CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        CalculateGaussianAdditionalWorkingBytes(source.Width, source.Height, sigma, false, out int radius);
        var weights = new double[radius * 2 + 1];
        double sum = 0D;
        for (int x = -radius; x <= radius; x++) { double w = Math.Exp(-x * x / (2D * sigma * sigma)); weights[x + radius] = w; sum += w; }
        for (int x = 0; x < weights.Length; x++) weights[x] /= sum;
        return Smooth(source, weights, cancellationToken);
    }

    /// <summary>Applies a Gaussian unsharp mask with unit strength while retaining the original alpha.</summary>
    public static OfficeRasterImage GaussianSharpen(OfficeRasterImage source, double sigma = 3D, CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        // Smooth owns float scratch as well as its output; the final color pass runs after that scratch is released.
        CalculateGaussianAdditionalWorkingBytes(source.Width, source.Height, sigma, true, out _);
        OfficeRasterImage blurred = GaussianBlur(source, sigma, cancellationToken);
        return Map(source, (c, x, y) => {
            OfficeColor soft = blurred.GetPixel(x, y);
            return c.A == 0 ? OfficeColor.FromRgba(0, 0, 0, 0) : Rgb(2D * c.R - soft.R, 2D * c.G - soft.G, 2D * c.B - soft.B, c.A);
        }, cancellationToken);
    }

    /// <summary>Applies a square box blur with a radius in pixels and premultiplied-alpha sampling.</summary>
    public static OfficeRasterImage BoxBlur(OfficeRasterImage source, int radius = 7, CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        CalculateBoxBlurAdditionalWorkingBytes(source.Width, source.Height, radius);
        var weights = new double[radius * 2 + 1];
        for (int i = 0; i < weights.Length; i++) weights[i] = 1D / weights.Length;
        return Smooth(source, weights, cancellationToken);
    }

    /// <summary>Applies a circular aperture blur in a gamma-adjusted channel space; radius is in pixels.</summary>
    /// <remarks>The disk aperture differs from Gaussian and square box kernels. The operation retains straight-alpha output.</remarks>
    public static OfficeRasterImage BokehBlur(OfficeRasterImage source, int radius = 5, double gamma = 3D, CancellationToken cancellationToken = default) {
        ValidateRadius(radius, 32, nameof(radius));
        ValidateAmount(gamma, nameof(gamma), .1D, 10D);
        if (source == null) throw new ArgumentNullException(nameof(source));
        int size = radius * 2 + 1;
        var kernel = new double[size * size];
        for (int y = -radius; y <= radius; y++) for (int x = -radius; x <= radius; x++) {
            if (x * x + y * y <= radius * radius) kernel[(y + radius) * size + x + radius] = 1D;
        }
        return ConvolveCore(source, kernel, size, size, true, gamma, cancellationToken, cloneKernel: false);
    }

    /// <summary>Applies an odd-sized row-major convolution kernel with clamped boundary samples and premultiplied alpha.</summary>
    /// <remarks>When normalize is true, coefficients are divided by their positive sum and alpha is filtered.
    /// Otherwise supplied coefficients are used directly and the source pixel's alpha is retained, supporting zero-sum edge kernels.</remarks>
    public static OfficeRasterImage Convolve(OfficeRasterImage source, double[] kernel, int kernelWidth, int kernelHeight,
        bool normalize = true, CancellationToken cancellationToken = default) =>
        ConvolveCore(source, kernel, kernelWidth, kernelHeight, normalize, 1D, cancellationToken);

    private static OfficeRasterImage ConvolveCore(OfficeRasterImage source, double[] kernel, int kernelWidth, int kernelHeight,
        bool normalize, double gamma, CancellationToken cancellationToken, bool cloneKernel = true) {
        cancellationToken.ThrowIfCancellationRequested();
        if (kernel == null) throw new ArgumentNullException(nameof(kernel));
        if (kernelWidth <= 0 || kernelWidth > 65 || (kernelWidth & 1) == 0) throw new ArgumentOutOfRangeException(nameof(kernelWidth));
        if (kernelHeight <= 0 || kernelHeight > 65 || (kernelHeight & 1) == 0) throw new ArgumentOutOfRangeException(nameof(kernelHeight));
        if (kernel.Length != kernelWidth * kernelHeight) throw new ArgumentException("Kernel dimensions do not match its coefficient count.", nameof(kernel));
        double sum = 0D;
        var weights = cloneKernel ? (double[])kernel.Clone() : kernel;
        foreach (double weight in weights) { ValidateAmount(weight, nameof(kernel), -1_000_000D, 1_000_000D); sum += weight; }
        if (normalize && sum <= 1E-12D) throw new ArgumentException("A normalized kernel must have a positive coefficient sum.", nameof(kernel));
        if (normalize) for (int i = 0; i < weights.Length; i++) weights[i] /= sum;
        ValidateSource(source, operationsPerPixel: weights.Length);
        var result = new OfficeRasterImage(source.Width, source.Height);
        byte[] input = source.PixelBuffer, output = result.PixelBuffer;
        int rx = kernelWidth / 2, ry = kernelHeight / 2;
        for (int y = 0; y < source.Height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int x = 0; x < source.Width; x++) {
                if ((x & 63) == 0) cancellationToken.ThrowIfCancellationRequested();
                double red = 0, green = 0, blue = 0, alpha = 0;
                for (int ky = 0; ky < kernelHeight; ky++) {
                    int sy = Math.Max(0, Math.Min(source.Height - 1, y + ky - ry));
                    for (int kx = 0; kx < kernelWidth; kx++) {
                        int sx = Math.Max(0, Math.Min(source.Width - 1, x + kx - rx));
                        int offset = (sy * source.Width + sx) * 4;
                        double w = weights[ky * kernelWidth + kx], a = input[offset + 3] / 255D * w;
                        red += GammaExpand(input[offset], gamma) * a;
                        green += GammaExpand(input[offset + 1], gamma) * a;
                        blue += GammaExpand(input[offset + 2], gamma) * a; alpha += a;
                    }
                }
                int target = (y * source.Width + x) * 4;
                if (!normalize) alpha = input[target + 3] / 255D;
                WritePremultiplied(output, target, red, green, blue, alpha, gamma);
            }
        }
        return result;
    }

    private static OfficeRasterImage Smooth(OfficeRasterImage source, double[] weights, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (source == null) throw new ArgumentNullException(nameof(source));
        CalculateSmoothingAdditionalWorkingBytes(source.Width, source.Height, weights.Length, false);
        float[] intermediate = new float[source.PixelBuffer.Length];
        var result = new OfficeRasterImage(source.Width, source.Height);
        byte[] input = source.PixelBuffer, output = result.PixelBuffer;
        int radius = weights.Length / 2;
        for (int y = 0; y < source.Height; y++) {
            token.ThrowIfCancellationRequested();
            for (int x = 0; x < source.Width; x++) {
                if ((x & 255) == 0) token.ThrowIfCancellationRequested();
                double r = 0, g = 0, b = 0, a = 0;
                for (int k = 0; k < weights.Length; k++) {
                    int sx = Math.Max(0, Math.Min(source.Width - 1, x + k - radius));
                    int offset = (y * source.Width + sx) * 4;
                    double alpha = input[offset + 3] / 255D * weights[k];
                    r += input[offset] * alpha; g += input[offset + 1] * alpha; b += input[offset + 2] * alpha; a += alpha;
                }
                int target = (y * source.Width + x) * 4;
                intermediate[target] = (float)r; intermediate[target + 1] = (float)g; intermediate[target + 2] = (float)b; intermediate[target + 3] = (float)a;
            }
        }
        for (int y = 0; y < source.Height; y++) {
            token.ThrowIfCancellationRequested();
            for (int x = 0; x < source.Width; x++) {
                if ((x & 255) == 0) token.ThrowIfCancellationRequested();
                double r = 0, g = 0, b = 0, a = 0;
                for (int k = 0; k < weights.Length; k++) {
                    int sy = Math.Max(0, Math.Min(source.Height - 1, y + k - radius));
                    int offset = (sy * source.Width + x) * 4;
                    double weight = weights[k];
                    r += intermediate[offset] * weight; g += intermediate[offset + 1] * weight; b += intermediate[offset + 2] * weight; a += intermediate[offset + 3] * weight;
                }
                WritePremultiplied(output, (y * source.Width + x) * 4, r, g, b, a);
            }
        }
        return result;
    }

    private static void WritePremultiplied(byte[] output, int offset, double r, double g, double b, double alpha, double gamma = 1D) {
        if (alpha <= 1E-12D) { output[offset] = output[offset + 1] = output[offset + 2] = output[offset + 3] = 0; return; }
        output[offset] = GammaCompress(r / alpha, gamma); output[offset + 1] = GammaCompress(g / alpha, gamma);
        output[offset + 2] = GammaCompress(b / alpha, gamma); output[offset + 3] = Channel(alpha * 255D);
    }

    private static double GammaExpand(byte value, double gamma) => gamma == 1D ? value : Math.Pow(value / 255D, gamma) * 255D;
    private static byte GammaCompress(double value, double gamma) => Channel(gamma == 1D ? value : Math.Pow(Math.Max(0D, value) / 255D, 1D / gamma) * 255D);

    private static void ValidateRadius(int value, int maximum, string name) {
        if (value < 1 || value > maximum) throw new ArgumentOutOfRangeException(name, value, "The pixel radius is outside the supported range.");
    }
}
