using System;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeRasterFilters {
    /// <summary>Replaces each square block with its premultiplied RGBA average; size is measured in pixels.</summary>
    public static OfficeRasterImage Pixelate(OfficeRasterImage source, int size = 4, CancellationToken cancellationToken = default) {
        ValidateRadius(size, 4096, nameof(size));
        cancellationToken.ThrowIfCancellationRequested();
        ValidateSource(source);
        var result = new OfficeRasterImage(source.Width, source.Height);
        for (int top = 0; top < source.Height; top += size) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int left = 0; left < source.Width; left += size) {
                int right = Math.Min(source.Width, left + size), bottom = Math.Min(source.Height, top + size);
                double r = 0, g = 0, b = 0, a = 0;
                for (int y = top; y < bottom; y++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    for (int x = left; x < right; x++) {
                        OfficeColor c = source.GetPixel(x, y);
                        double alpha = c.A / 255D;
                        r += c.R * alpha; g += c.G * alpha; b += c.B * alpha; a += alpha;
                    }
                }
                int count = (right - left) * (bottom - top);
                OfficeColor average = a <= 0D ? OfficeColor.FromRgba(0, 0, 0, 0) : Rgb(r / a, g / a, b / a, Channel(a * 255D / count));
                for (int y = top; y < bottom; y++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    for (int x = left; x < right; x++) result.SetPixel(x, y, average);
                }
            }
        }
        return result;
    }

    /// <summary>Chooses the alpha-weighted dominant luminance bucket in a square brush, retaining each source pixel's alpha.</summary>
    /// <remarks>Levels is the number of luminance buckets; brushSize is the square window width in pixels.</remarks>
    public static OfficeRasterImage OilPaint(OfficeRasterImage source, int levels = 10, int brushSize = 15, CancellationToken cancellationToken = default) {
        ValidateRadius(levels, 256, nameof(levels)); ValidateRadius(brushSize, 65, nameof(brushSize));
        cancellationToken.ThrowIfCancellationRequested();
        ValidateSource(source, levels * 4L * sizeof(double), (long)brushSize * brushSize + levels);
        var result = new OfficeRasterImage(source.Width, source.Height);
        var counts = new double[levels]; var red = new double[levels]; var green = new double[levels]; var blue = new double[levels];
        int before = brushSize / 2, after = brushSize - before - 1;
        for (int y = 0; y < source.Height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int x = 0; x < source.Width; x++) {
                if ((x & 31) == 0) cancellationToken.ThrowIfCancellationRequested();
                Array.Clear(counts, 0, levels); Array.Clear(red, 0, levels); Array.Clear(green, 0, levels); Array.Clear(blue, 0, levels);
                for (int sy = Math.Max(0, y - before); sy <= Math.Min(source.Height - 1, y + after); sy++) {
                    for (int sx = Math.Max(0, x - before); sx <= Math.Min(source.Width - 1, x + after); sx++) {
                        OfficeColor c = source.GetPixel(sx, sy);
                        int bucket = Math.Min(levels - 1, (int)(Luminance(c) * levels / 256D));
                        double alpha = c.A / 255D;
                        counts[bucket] += alpha; red[bucket] += c.R * alpha; green[bucket] += c.G * alpha; blue[bucket] += c.B * alpha;
                    }
                }
                int strongest = 0;
                for (int i = 1; i < levels; i++) if (counts[i] > counts[strongest]) strongest = i;
                OfficeColor center = source.GetPixel(x, y);
                result.SetPixel(x, y, counts[strongest] <= 0D ? OfficeColor.FromRgba(0, 0, 0, center.A) :
                    Rgb(red[strongest] / counts[strongest], green[strongest] / counts[strongest], blue[strongest] / counts[strongest], center.A));
            }
        }
        return result;
    }

    /// <summary>Applies an alpha-weighted local-mean threshold using a pixel radius and relative contrast between zero and one.</summary>
    public static OfficeRasterImage AdaptiveThreshold(OfficeRasterImage source, int radius = 15, double contrast = .15D, CancellationToken cancellationToken = default) {
        ValidateRadius(radius, 4096, nameof(radius)); ValidateAmount(contrast, nameof(contrast), maximum: 1D);
        cancellationToken.ThrowIfCancellationRequested();
        if (source == null) throw new ArgumentNullException(nameof(source));
        long entries = checked(((long)source.Width + 1L) * ((long)source.Height + 1L));
        ValidateSource(source, checked(entries * sizeof(double) * 2L));
        int stride = checked(source.Width + 1);
        var sums = new double[checked((int)entries)];
        var alphaSums = new double[checked((int)entries)];
        for (int y = 0; y < source.Height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            double row = 0D, alphaRow = 0D;
            for (int x = 0; x < source.Width; x++) {
                if ((x & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                OfficeColor c = source.GetPixel(x, y); double a = c.A / 255D;
                row += Luminance(c) * a; alphaRow += a;
                int index = (y + 1) * stride + x + 1;
                sums[index] = sums[index - stride] + row; alphaSums[index] = alphaSums[index - stride] + alphaRow;
            }
        }
        return Map(source, (c, x, y) => {
            int left = Math.Max(0, x - radius), right = Math.Min(source.Width, x + radius + 1);
            int top = Math.Max(0, y - radius), bottom = Math.Min(source.Height, y + radius + 1);
            double a = RegionSum(alphaSums, stride, left, top, right, bottom);
            double mean = a <= 0D ? 0D : RegionSum(sums, stride, left, top, right, bottom) / a;
            byte channel = Luminance(c) >= mean * (1D - contrast) ? (byte)255 : (byte)0;
            return OfficeColor.FromRgba(channel, channel, channel, c.A);
        }, cancellationToken);
    }

    /// <summary>Applies alpha-weighted luminance histogram equalization while retaining alpha and approximate RGB hue.</summary>
    public static OfficeRasterImage HistogramEqualization(OfficeRasterImage source, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested(); ValidateSource(source);
        var histogram = new double[256];
        byte[] input = source.PixelBuffer;
        for (int offset = 0; offset < input.Length; offset += 4) {
            if ((offset & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            int luminance = Channel(input[offset] * .2126D + input[offset + 1] * .7152D + input[offset + 2] * .0722D);
            histogram[luminance] += input[offset + 3] / 255D;
        }
        var values = new double[256]; double total = 0D, first = 0D;
        for (int i = 0; i < 256; i++) {
            total += histogram[i]; values[i] = total;
            if (first == 0D && histogram[i] > 0D) first = total;
        }
        if (total <= first) return source.Clone();
        for (int i = 0; i < 256; i++) values[i] = Math.Max(0D, (values[i] - first) / (total - first) * 255D);
        return Map(source, (c, _, _) => {
            double old = Luminance(c), target = values[Channel(old)];
            return old <= 1E-12D ? Rgb(target, target, target, c.A) : Rgb(c.R * target / old, c.G * target / old, c.B * target / old, c.A);
        }, cancellationToken);
    }

    /// <summary>Produces ordered black-and-white dithering with a four-by-four Bayer matrix, retaining source alpha.</summary>
    public static OfficeRasterImage Dither(OfficeRasterImage source, CancellationToken cancellationToken = default) {
        int[] bayer = { 0, 8, 2, 10, 12, 4, 14, 6, 3, 11, 1, 9, 15, 7, 13, 5 };
        return Map(source, (c, x, y) => {
            byte level = Luminance(c) / 255D > (bayer[(y & 3) * 4 + (x & 3)] + .5D) / 16D ? (byte)255 : (byte)0;
            return OfficeColor.FromRgba(level, level, level, c.A);
        }, cancellationToken);
    }

    private static double RegionSum(double[] values, int stride, int left, int top, int right, int bottom) =>
        values[bottom * stride + right] - values[top * stride + right] - values[bottom * stride + left] + values[top * stride + left];
}
