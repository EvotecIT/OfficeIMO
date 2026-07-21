using System;

namespace OfficeIMO.Drawing;

internal static class OfficeRasterGuards {
    internal const long MaximumPixels = 50_000_000L;
    internal const int MaximumEncodedBytes = 128 * 1024 * 1024;

    public static void EnsurePayloadWithinLimits(int length, string message) {
        if (length < 0 || length > MaximumEncodedBytes) throw new FormatException(message);
    }

    public static bool TryEnsurePixelCount(int width, int height, out int pixels) =>
        TryEnsurePixelCount(width, height, MaximumPixels, out pixels);

    public static bool TryEnsurePixelCount(int width, int height, long maximumPixels, out int pixels) {
        pixels = 0;
        if (width <= 0 || height <= 0) return false;
        long total = (long)width * height;
        if (total > int.MaxValue || (maximumPixels > 0 && total > maximumPixels)) return false;
        pixels = (int)total;
        return true;
    }

    public static int EnsurePixelCount(int width, int height, string message) {
        if (TryEnsurePixelCount(width, height, out int pixels)) return pixels;
        throw new FormatException(message);
    }

    public static int EnsureByteCount(long bytes, string message) {
        if (bytes <= 0 || bytes > int.MaxValue || bytes > MaximumEncodedBytes) throw new FormatException(message);
        return (int)bytes;
    }

    public static byte[] AllocatePixelBuffer(int width, int height, string message) =>
        new byte[EnsurePixelCount(width, height, message)];

    public static byte[] AllocateRgba32(int width, int height, string message) {
        int pixels = EnsurePixelCount(width, height, message);
        return new byte[checked(pixels * 4)];
    }

    public static int EnsureOutputPixels(int width, int height, string message) {
        if (TryEnsurePixelCount(width, height, out int pixels)) return pixels;
        throw new ArgumentException(message);
    }

    public static int EnsureOutputBytes(long bytes, string message) {
        if (bytes <= 0 || bytes > int.MaxValue || bytes > MaximumEncodedBytes) throw new ArgumentException(message);
        return (int)bytes;
    }
}
