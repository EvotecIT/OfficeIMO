using System;

namespace OfficeIMO.Drawing;

/// <summary>Sampling algorithm used when resizing dependency-free raster images.</summary>
public enum OfficeRasterResamplingMode {
    /// <summary>Chooses the closest source pixel and preserves hard pixel edges.</summary>
    NearestNeighbor,
    /// <summary>Interpolates four source pixels in premultiplied-alpha space.</summary>
    Bilinear
}

/// <summary>Dependency-free RGBA image resampling shared by document renderers.</summary>
public static class OfficeRasterResampler {
    /// <summary>Resizes an RGBA image to exact pixel dimensions.</summary>
    public static OfficeRasterImage Resize(OfficeRasterImage source, int width, int height, OfficeRasterResamplingMode mode = OfficeRasterResamplingMode.Bilinear) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (width <= 0) throw new ArgumentOutOfRangeException(nameof(width));
        if (height <= 0) throw new ArgumentOutOfRangeException(nameof(height));
        if (mode != OfficeRasterResamplingMode.NearestNeighbor && mode != OfficeRasterResamplingMode.Bilinear) throw new ArgumentOutOfRangeException(nameof(mode));
        OfficeRasterGuards.EnsureOutputPixels(width, height, "Raster resize dimensions exceed the managed image limit.");

        if (source.Width == width && source.Height == height) {
            return OfficeRasterImage.FromRgba32(width, height, source.GetPixels());
        }

        var result = new OfficeRasterImage(width, height);
        byte[] input = source.PixelBuffer;
        byte[] output = result.PixelBuffer;
        double scaleX = source.Width / (double)width;
        double scaleY = source.Height / (double)height;
        for (int y = 0; y < height; y++) {
            double sourceY = ((y + 0.5D) * scaleY) - 0.5D;
            for (int x = 0; x < width; x++) {
                double sourceX = ((x + 0.5D) * scaleX) - 0.5D;
                int target = ((y * width) + x) * 4;
                if (mode == OfficeRasterResamplingMode.NearestNeighbor) {
                    CopyNearest(input, source.Width, source.Height, sourceX, sourceY, output, target);
                } else {
                    CopyBilinear(input, source.Width, source.Height, sourceX, sourceY, output, target);
                }
            }
        }

        return result;
    }

    private static void CopyNearest(byte[] input, int width, int height, double x, double y, byte[] output, int target) {
        int sourceX = Clamp((int)Math.Floor(x + 0.5D), 0, width - 1);
        int sourceY = Clamp((int)Math.Floor(y + 0.5D), 0, height - 1);
        int source = ((sourceY * width) + sourceX) * 4;
        output[target] = input[source];
        output[target + 1] = input[source + 1];
        output[target + 2] = input[source + 2];
        output[target + 3] = input[source + 3];
    }

    private static void CopyBilinear(byte[] input, int width, int height, double x, double y, byte[] output, int target) {
        double sampleX = Clamp(x, 0D, width - 1D);
        double sampleY = Clamp(y, 0D, height - 1D);
        int x0 = (int)Math.Floor(sampleX);
        int y0 = (int)Math.Floor(sampleY);
        int x1 = Clamp(x0 + 1, 0, width - 1);
        int y1 = Clamp(y0 + 1, 0, height - 1);
        double tx = sampleX - x0;
        double ty = sampleY - y0;
        double w00 = (1D - tx) * (1D - ty);
        double w10 = tx * (1D - ty);
        double w01 = (1D - tx) * ty;
        double w11 = tx * ty;
        int p00 = ((y0 * width) + x0) * 4;
        int p10 = ((y0 * width) + x1) * 4;
        int p01 = ((y1 * width) + x0) * 4;
        int p11 = ((y1 * width) + x1) * 4;
        double alpha = (input[p00 + 3] * w00) + (input[p10 + 3] * w10) + (input[p01 + 3] * w01) + (input[p11 + 3] * w11);
        if (alpha <= 0D) {
            output[target] = output[target + 1] = output[target + 2] = output[target + 3] = 0;
            return;
        }

        output[target] = InterpolateChannel(input, p00, p10, p01, p11, 0, w00, w10, w01, w11, alpha);
        output[target + 1] = InterpolateChannel(input, p00, p10, p01, p11, 1, w00, w10, w01, w11, alpha);
        output[target + 2] = InterpolateChannel(input, p00, p10, p01, p11, 2, w00, w10, w01, w11, alpha);
        output[target + 3] = (byte)Math.Round(Clamp(alpha, 0D, 255D));
    }

    private static byte InterpolateChannel(byte[] pixels, int p00, int p10, int p01, int p11, int channel, double w00, double w10, double w01, double w11, double alpha) {
        double value = (pixels[p00 + channel] * pixels[p00 + 3] * w00) +
            (pixels[p10 + channel] * pixels[p10 + 3] * w10) +
            (pixels[p01 + channel] * pixels[p01 + 3] * w01) +
            (pixels[p11 + channel] * pixels[p11 + 3] * w11);
        return (byte)Math.Round(Clamp(value / alpha, 0D, 255D));
    }

    private static int Clamp(int value, int minimum, int maximum) => value < minimum ? minimum : value > maximum ? maximum : value;
    private static double Clamp(double value, double minimum, double maximum) => value < minimum ? minimum : value > maximum ? maximum : value;
}
