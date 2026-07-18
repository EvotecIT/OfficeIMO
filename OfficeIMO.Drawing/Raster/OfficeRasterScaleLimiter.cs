using System;

namespace OfficeIMO.Drawing;

/// <summary>Resolved raster dimensions after applying a hard decoded-pixel limit.</summary>
public readonly struct OfficeRasterScaleLimit {
    internal OfficeRasterScaleLimit(double scale, int pixelWidth, int pixelHeight, bool wasLimited) {
        Scale = scale;
        PixelWidth = pixelWidth;
        PixelHeight = pixelHeight;
        WasLimited = wasLimited;
    }

    /// <summary>Effective scale that satisfies the requested ceiling.</summary>
    public double Scale { get; }
    /// <summary>Ceiling-rounded output width.</summary>
    public int PixelWidth { get; }
    /// <summary>Ceiling-rounded output height.</summary>
    public int PixelHeight { get; }
    /// <summary>Whether the requested scale had to be reduced.</summary>
    public bool WasLimited { get; }
    /// <summary>Total decoded pixels.</summary>
    public long PixelCount => (long)PixelWidth * PixelHeight;
}

/// <summary>Shared overflow-safe raster scale limiter for Office document engines.</summary>
public static class OfficeRasterScaleLimiter {
    /// <summary>Resolves the largest scale no greater than <paramref name="requestedScale"/> whose rounded dimensions fit the pixel ceiling.</summary>
    public static OfficeRasterScaleLimit Resolve(double width, double height, double requestedScale, long maximumPixels) =>
        Resolve(width, height, requestedScale, maximumPixels, int.MaxValue);

    /// <summary>Resolves the largest scale whose rounded dimensions fit both the decoded-pixel and per-dimension ceilings.</summary>
    public static OfficeRasterScaleLimit Resolve(double width, double height, double requestedScale, long maximumPixels, int maximumDimension) {
        ValidatePositive(width, nameof(width));
        ValidatePositive(height, nameof(height));
        ValidatePositive(requestedScale, nameof(requestedScale));
        if (maximumPixels < 1L) throw new ArgumentOutOfRangeException(nameof(maximumPixels));
        if (maximumDimension < 1) throw new ArgumentOutOfRangeException(nameof(maximumDimension));

        if (Fits(width, height, requestedScale, maximumPixels, maximumDimension, out int requestedWidth, out int requestedHeight)) {
            return new OfficeRasterScaleLimit(requestedScale, requestedWidth, requestedHeight, false);
        }

        double exponent = (Math.Log(maximumPixels) - Math.Log(width) - Math.Log(height)) / 2D;
        double upper = Math.Min(requestedScale, Math.Exp(exponent));
        upper = Math.Min(upper, maximumDimension / width);
        upper = Math.Min(upper, maximumDimension / height);
        if (double.IsNaN(upper) || upper <= 0D) upper = Math.Min(requestedScale, double.Epsilon);
        double limited = upper;
        if (!Fits(width, height, limited, maximumPixels, maximumDimension, out _, out _)) {
            double lower = 0D;
            for (int iteration = 0; iteration < 80; iteration++) {
                double candidate = lower + (upper - lower) / 2D;
                if (Fits(width, height, candidate, maximumPixels, maximumDimension, out _, out _)) lower = candidate;
                else upper = candidate;
            }
            limited = lower > 0D ? lower : double.Epsilon;
        }

        if (!Fits(width, height, limited, maximumPixels, maximumDimension, out int pixelWidth, out int pixelHeight)) {
            throw new InvalidOperationException("The raster scale could not be reduced to the requested decoded-raster limits.");
        }
        return new OfficeRasterScaleLimit(limited, pixelWidth, pixelHeight, true);
    }

    private static bool Fits(double width, double height, double scale, long maximumPixels, int maximumDimension, out int pixelWidth, out int pixelHeight) {
        pixelWidth = 0;
        pixelHeight = 0;
        if (scale < 0D || double.IsNaN(scale) || double.IsInfinity(scale)) return false;
        double scaledWidth = Math.Max(1D, Math.Ceiling(width * scale));
        double scaledHeight = Math.Max(1D, Math.Ceiling(height * scale));
        if (double.IsNaN(scaledWidth) || double.IsInfinity(scaledWidth) || scaledWidth > int.MaxValue ||
            double.IsNaN(scaledHeight) || double.IsInfinity(scaledHeight) || scaledHeight > int.MaxValue) return false;
        pixelWidth = (int)scaledWidth;
        pixelHeight = (int)scaledHeight;
        return pixelWidth <= maximumDimension && pixelHeight <= maximumDimension && pixelWidth <= maximumPixels / pixelHeight;
    }

    private static void ValidatePositive(double value, string name) {
        if (value <= 0D || double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(name);
    }
}
