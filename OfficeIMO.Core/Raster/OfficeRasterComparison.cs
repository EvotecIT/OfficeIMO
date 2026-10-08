using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Pixel comparison evidence for equal-size rasters in premultiplied RGBA space.</summary>
public sealed class OfficeRasterComparisonResult {
    internal OfficeRasterComparisonResult(OfficeRasterImage difference, double mean, long changed, byte maximum) {
        DifferenceImage = difference; MeanAbsoluteDifference = mean; ChangedPixels = changed; MaximumChannelDifference = maximum;
    }
    /// <summary>Opaque RGB visualization of the channel differences; alpha differences are added to each visible channel.</summary>
    public OfficeRasterImage DifferenceImage { get; }
    /// <summary>Mean absolute premultiplied RGBA difference, normalized between zero and one.</summary>
    public double MeanAbsoluteDifference { get; }
    /// <summary>One minus the normalized mean absolute difference.</summary>
    public double Similarity => 1D - MeanAbsoluteDifference;
    /// <summary>Number of pixels with a difference in at least one premultiplied RGBA channel.</summary>
    public long ChangedPixels { get; }
    /// <summary>Largest absolute premultiplied channel difference in byte units.</summary>
    public byte MaximumChannelDifference { get; }
}

/// <summary>Dependency-free, alpha-aware comparison of equal-sized raster images.</summary>
public static class OfficeRasterComparison {
    /// <summary>Compares premultiplied RGB and alpha, ignoring invisible RGB in completely transparent pixels.</summary>
    public static OfficeRasterComparisonResult Compare(OfficeRasterImage left, OfficeRasterImage right, CancellationToken cancellationToken = default) {
        if (left == null) throw new ArgumentNullException(nameof(left));
        if (right == null) throw new ArgumentNullException(nameof(right));
        cancellationToken.ThrowIfCancellationRequested();
        if (left.Width != right.Width || left.Height != right.Height) throw new ArgumentException("Compared raster dimensions must match.", nameof(right));
        long count = OfficeRasterGuards.EnsureOutputPixels(left.Width, left.Height, "Raster comparison dimensions exceed the managed image limit.");
        if (count * 12L + 64L * 1024L > OfficeRasterGuards.MaximumDecodedBytes) throw new ArgumentException("Raster comparison working set exceeds the managed image limit.", nameof(left));
        var difference = new OfficeRasterImage(left.Width, left.Height);
        byte[] a = left.PixelBuffer, b = right.PixelBuffer, output = difference.PixelBuffer;
        double sum = 0D; long changed = 0L; byte maximum = 0;
        for (int offset = 0; offset < a.Length; offset += 4) {
            if ((offset & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            double da = Math.Abs(a[offset + 3] - b[offset + 3]);
            double dr = Math.Abs(a[offset] * a[offset + 3] / 255D - b[offset] * b[offset + 3] / 255D);
            double dg = Math.Abs(a[offset + 1] * a[offset + 3] / 255D - b[offset + 1] * b[offset + 3] / 255D);
            double db = Math.Abs(a[offset + 2] * a[offset + 3] / 255D - b[offset + 2] * b[offset + 3] / 255D);
            sum += da + dr + dg + db;
            if (da + dr + dg + db > 0D) changed++;
            maximum = (byte)Math.Max(maximum, Math.Ceiling(Math.Max(da, Math.Max(dr, Math.Max(dg, db)))));
            output[offset] = (byte)Math.Min(255D, Math.Round(dr + da));
            output[offset + 1] = (byte)Math.Min(255D, Math.Round(dg + da));
            output[offset + 2] = (byte)Math.Min(255D, Math.Round(db + da));
            output[offset + 3] = 255;
        }
        return new OfficeRasterComparisonResult(difference, sum / (count * 1020D), changed, maximum);
    }
}
