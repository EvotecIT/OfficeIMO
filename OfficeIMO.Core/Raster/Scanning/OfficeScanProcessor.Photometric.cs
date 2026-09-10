using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeScanProcessor {
    private static void ApplyPhotometricCorrections(OfficeRasterImage image, OfficeScanProcessingOptions options,
        List<OfficeScanProcessingStep> steps, CancellationToken token) {
        byte[]? background = options.NormalizeBackground ? EstimateBackground(image, options.BackgroundRadius, token) : null;
        byte[] pixels = image.PixelBuffer;
        bool applyLevels = options.BlackPoint != 0 || options.WhitePoint != 255 || options.Gamma != 1D;
        var levels = new byte[256];
        for (int value = 0; value < levels.Length; value++) {
            double normalized = Math.Max(0D, Math.Min(1D, (value - options.BlackPoint) / (double)(options.WhitePoint - options.BlackPoint)));
            levels[value] = (byte)Math.Round(255D * Math.Pow(normalized, 1D / options.Gamma));
        }
        for (int y = 0; y < image.Height; y++) {
            token.ThrowIfCancellationRequested();
            for (int x = 0; x < image.Width; x++) {
                if ((x & 1023) == 0) token.ThrowIfCancellationRequested();
                int index = y * image.Width + x, offset = index * 4;
                int alpha = pixels[offset + 3];
                int r = CompositeWhite(pixels[offset], alpha), g = CompositeWhite(pixels[offset + 1], alpha), b = CompositeWhite(pixels[offset + 2], alpha);
                if (background != null && background[index] > 0) {
                    int paper = background[index];
                    r = Math.Min(255, (r * 255 + paper / 2) / paper);
                    g = Math.Min(255, (g * 255 + paper / 2) / paper);
                    b = Math.Min(255, (b * 255 + paper / 2) / paper);
                }
                if (options.ColorMode != OfficeScanColorMode.PreserveColor) r = g = b = (r * 77 + g * 150 + b * 29 + 128) >> 8;
                if (applyLevels) { r = levels[r]; g = levels[g]; b = levels[b]; }
                pixels[offset] = (byte)r; pixels[offset + 1] = (byte)g; pixels[offset + 2] = (byte)b; pixels[offset + 3] = 255;
            }
        }
        if (options.ColorMode == OfficeScanColorMode.Bilevel) ApplyBilevel(image, options.BilevelThreshold, token);
        steps.Add(new OfficeScanProcessingStep("background", options.NormalizeBackground,
            options.NormalizeBackground ? "Normalized local paper brightness; composited transparency over white." : "Background normalization was disabled; transparency was composited over white."));
        steps.Add(new OfficeScanProcessingStep("color", options.ColorMode != OfficeScanColorMode.PreserveColor,
            "Output color mode: " + options.ColorMode + "."));
        steps.Add(new OfficeScanProcessingStep("levels", applyLevels,
            applyLevels ? "Applied black point, white point, and midtone gamma." : "Retained the input tonal range."));
    }

    private static byte[] EstimateBackground(OfficeRasterImage image, int radius, CancellationToken token) {
        int width = image.Width, height = image.Height;
        var horizontal = new byte[checked(width * height)];
        var background = new byte[horizontal.Length];
        var deque = new int[Math.Max(width, height)];
        for (int y = 0; y < height; y++) {
            token.ThrowIfCancellationRequested();
            int head = 0, tail = 0, next = 0;
            for (int x = 0; x < width; x++) {
                if ((x & 1023) == 0) token.ThrowIfCancellationRequested();
                while (next < width && next <= x + radius) {
                    byte value = Luminance(image.PixelBuffer, (y * width + next) * 4);
                    while (tail > head && Luminance(image.PixelBuffer, (y * width + deque[tail - 1]) * 4) <= value) tail--;
                    deque[tail++] = next++;
                }
                while (head < tail && deque[head] < x - radius) head++;
                horizontal[y * width + x] = Luminance(image.PixelBuffer, (y * width + deque[head]) * 4);
            }
        }
        for (int x = 0; x < width; x++) {
            token.ThrowIfCancellationRequested();
            int head = 0, tail = 0, next = 0;
            for (int y = 0; y < height; y++) {
                if ((y & 1023) == 0) token.ThrowIfCancellationRequested();
                while (next < height && next <= y + radius) {
                    byte value = horizontal[next * width + x];
                    while (tail > head && horizontal[deque[tail - 1] * width + x] <= value) tail--;
                    deque[tail++] = next++;
                }
                while (head < tail && deque[head] < y - radius) head++;
                background[y * width + x] = horizontal[deque[head] * width + x];
            }
        }
        return background;
    }

    private static void ApplyBilevel(OfficeRasterImage image, int? threshold, CancellationToken token) {
        int[] histogram = Histogram(image, token);
        int selected = threshold ?? HistogramThreshold(histogram, image.Width * image.Height);
        byte[] pixels = image.PixelBuffer;
        for (int i = 0; i < pixels.Length; i += 4) {
            if ((i & 4095) == 0) token.ThrowIfCancellationRequested();
            byte value = Luminance(pixels, i) <= selected ? (byte)0 : (byte)255;
            pixels[i] = pixels[i + 1] = pixels[i + 2] = value; pixels[i + 3] = 255;
        }
    }

    private static (double Foreground, bool Blank) MeasureForeground(OfficeRasterImage image, CancellationToken token) {
        int[] histogram = Histogram(image, token);
        int total = image.Width * image.Height;
        int threshold = Math.Min(239, HistogramThreshold(histogram, total));
        long foreground = 0, dark = 0;
        for (int value = 0; value < 240; value++) { dark += histogram[value]; if (value <= threshold) foreground += histogram[value]; }
        return (foreground / (double)total, dark == 0);
    }

    private static int[] Histogram(OfficeRasterImage image, CancellationToken token) {
        var histogram = new int[256];
        for (int i = 0; i < image.PixelBuffer.Length; i += 4) {
            if ((i & 4095) == 0) token.ThrowIfCancellationRequested();
            histogram[Luminance(image.PixelBuffer, i)]++;
        }
        return histogram;
    }

    // Maximizes between-class variance of a gray histogram (Otsu thresholding).
    private static int HistogramThreshold(int[] histogram, int total) {
        long sum = 0, leftSum = 0, leftCount = 0;
        for (int i = 0; i < 256; i++) sum += (long)i * histogram[i];
        double best = 0D;
        int threshold = 127;
        for (int i = 0; i < 255; i++) {
            leftCount += histogram[i]; leftSum += (long)i * histogram[i];
            long rightCount = total - leftCount;
            if (leftCount == 0 || rightCount == 0) continue;
            double difference = leftSum / (double)leftCount - (sum - leftSum) / (double)rightCount;
            double variance = leftCount * (double)rightCount * difference * difference;
            if (variance > best) { best = variance; threshold = i; }
        }
        return threshold;
    }

    private static byte Luminance(byte[] pixels, int offset) {
        int gray = (pixels[offset] * 77 + pixels[offset + 1] * 150 + pixels[offset + 2] * 29 + 128) >> 8;
        return (byte)CompositeWhite(gray, pixels[offset + 3]);
    }
    private static int CompositeWhite(int channel, int alpha) => (channel * alpha + 255 * (255 - alpha) + 127) / 255;
}