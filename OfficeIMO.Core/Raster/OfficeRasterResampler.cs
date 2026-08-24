using System;

namespace OfficeIMO.Drawing;

/// <summary>Sampling algorithm used when resizing dependency-free raster images.</summary>
public enum OfficeRasterResamplingMode {
    /// <summary>Chooses the closest source pixel and preserves hard pixel edges.</summary>
    NearestNeighbor,
    /// <summary>Interpolates four source pixels in premultiplied-alpha space.</summary>
    Bilinear,
    /// <summary>Area-averages downsampled axes and linearly interpolates enlarged axes in premultiplied-alpha space.</summary>
    Area,
    /// <summary>Uses a radius-three Lanczos filter with antialiasing and premultiplied-alpha sampling.</summary>
    Lanczos3
}

/// <summary>Color space in which raster color channels are filtered.</summary>
public enum OfficeRasterResamplingColorSpace {
    /// <summary>Filters encoded sRGB channel values for compatibility with existing output.</summary>
    EncodedSrgb,
    /// <summary>Converts sRGB channels to linear light before filtering and encodes the result back to sRGB.</summary>
    LinearLight
}

/// <summary>Dependency-free RGBA image resampling shared by document renderers.</summary>
public static partial class OfficeRasterResampler {
    /// <summary>Resizes an RGBA image to exact pixel dimensions.</summary>
    public static OfficeRasterImage Resize(OfficeRasterImage source, int width, int height, OfficeRasterResamplingMode mode = OfficeRasterResamplingMode.Bilinear) =>
        Resize(source, width, height, mode, OfficeRasterResamplingColorSpace.EncodedSrgb);

    /// <summary>Resizes an RGBA image to exact pixel dimensions using an explicit filtering color space.</summary>
    public static OfficeRasterImage Resize(
        OfficeRasterImage source,
        int width,
        int height,
        OfficeRasterResamplingMode mode,
        OfficeRasterResamplingColorSpace colorSpace) =>
        Resize(source, width, height, mode, colorSpace, retainedManagedBytes: 0L);

    internal static OfficeRasterImage Resize(
        OfficeRasterImage source,
        int width,
        int height,
        OfficeRasterResamplingMode mode,
        OfficeRasterResamplingColorSpace colorSpace,
        long retainedManagedBytes) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (width <= 0) throw new ArgumentOutOfRangeException(nameof(width));
        if (height <= 0) throw new ArgumentOutOfRangeException(nameof(height));
        if (retainedManagedBytes < 0L) throw new ArgumentOutOfRangeException(nameof(retainedManagedBytes));
        if (mode != OfficeRasterResamplingMode.NearestNeighbor &&
            mode != OfficeRasterResamplingMode.Bilinear &&
            mode != OfficeRasterResamplingMode.Area &&
            mode != OfficeRasterResamplingMode.Lanczos3) {
            throw new ArgumentOutOfRangeException(nameof(mode));
        }
        if (colorSpace != OfficeRasterResamplingColorSpace.EncodedSrgb &&
            colorSpace != OfficeRasterResamplingColorSpace.LinearLight) {
            throw new ArgumentOutOfRangeException(nameof(colorSpace));
        }
        OfficeRasterGuards.EnsureOutputPixels(width, height, "Raster resize dimensions exceed the managed image limit.");

        if (source.Width == width && source.Height == height) {
            EnsureSimpleWorkingSet(source, width, height, retainedManagedBytes);
            return OfficeRasterImage.FromRgba32(width, height, source.PixelBuffer);
        }

        if (mode == OfficeRasterResamplingMode.Area || mode == OfficeRasterResamplingMode.Lanczos3) {
            return ResizeSeparable(source, width, height, mode, colorSpace, retainedManagedBytes);
        }

        EnsureSimpleWorkingSet(source, width, height, retainedManagedBytes);
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
                    CopyBilinear(input, source.Width, source.Height, sourceX, sourceY, output, target, colorSpace);
                }
            }
        }

        return result;
    }

    internal static bool TryGetSimpleWorkingSetBytes(
        int sourceWidth,
        int sourceHeight,
        int width,
        int height,
        long retainedManagedBytes,
        out long workingSetBytes) {
        workingSetBytes = 0L;
        if (sourceWidth <= 0 || sourceHeight <= 0 || width <= 0 || height <= 0 || retainedManagedBytes < 0L) {
            return false;
        }
        try {
            workingSetBytes = checked(
                (long)sourceWidth * sourceHeight * 4L + 24L +
                (long)width * height * 4L + 24L +
                retainedManagedBytes + 64L * 1024L);
            return workingSetBytes <= OfficeRasterGuards.MaximumDecodedBytes;
        } catch (OverflowException) {
            return false;
        }
    }

    private static void EnsureSimpleWorkingSet(
        OfficeRasterImage source,
        int width,
        int height,
        long retainedManagedBytes) {
        if (!TryGetSimpleWorkingSetBytes(
                source.Width, source.Height, width, height, retainedManagedBytes, out _)) {
            throw new ArgumentException(
                "Raster resampling working set exceeds the managed image limit.", nameof(source));
        }
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

    private static void CopyBilinear(
        byte[] input,
        int width,
        int height,
        double x,
        double y,
        byte[] output,
        int target,
        OfficeRasterResamplingColorSpace colorSpace) {
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

        output[target] = InterpolateChannel(input, p00, p10, p01, p11, 0, w00, w10, w01, w11, alpha, colorSpace);
        output[target + 1] = InterpolateChannel(input, p00, p10, p01, p11, 1, w00, w10, w01, w11, alpha, colorSpace);
        output[target + 2] = InterpolateChannel(input, p00, p10, p01, p11, 2, w00, w10, w01, w11, alpha, colorSpace);
        output[target + 3] = (byte)Math.Round(Clamp(alpha, 0D, 255D));
    }

    private static byte InterpolateChannel(
        byte[] pixels,
        int p00,
        int p10,
        int p01,
        int p11,
        int channel,
        double w00,
        double w10,
        double w01,
        double w11,
        double alpha,
        OfficeRasterResamplingColorSpace colorSpace) {
        double value = (DecodeChannel(pixels[p00 + channel], colorSpace) * pixels[p00 + 3] * w00) +
            (DecodeChannel(pixels[p10 + channel], colorSpace) * pixels[p10 + 3] * w10) +
            (DecodeChannel(pixels[p01 + channel], colorSpace) * pixels[p01 + 3] * w01) +
            (DecodeChannel(pixels[p11 + channel], colorSpace) * pixels[p11 + 3] * w11);
        return EncodeChannel(value / alpha, colorSpace);
    }

    private static double DecodeChannel(byte value, OfficeRasterResamplingColorSpace colorSpace) {
        if (colorSpace == OfficeRasterResamplingColorSpace.EncodedSrgb) return value;
        double encoded = value / 255D;
        double linear = encoded <= 0.04045D
            ? encoded / 12.92D
            : Math.Pow((encoded + 0.055D) / 1.055D, 2.4D);
        return linear * 255D;
    }

    private static byte EncodeChannel(double value, OfficeRasterResamplingColorSpace colorSpace) {
        value = Clamp(value, 0D, 255D);
        if (colorSpace == OfficeRasterResamplingColorSpace.EncodedSrgb) return (byte)Math.Round(value);
        double linear = value / 255D;
        double encoded = linear <= 0.0031308D
            ? linear * 12.92D
            : 1.055D * Math.Pow(linear, 1D / 2.4D) - 0.055D;
        return (byte)Math.Round(Clamp(encoded * 255D, 0D, 255D));
    }

    private static int Clamp(int value, int minimum, int maximum) => value < minimum ? minimum : value > maximum ? maximum : value;
    private static double Clamp(double value, double minimum, double maximum) => value < minimum ? minimum : value > maximum ? maximum : value;
}
