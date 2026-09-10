using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Separately owned perspective-corrected image and its exact coordinate mapping.</summary>
public sealed class OfficeScanPerspectiveResult {
    internal OfficeScanPerspectiveResult(OfficeRasterImage image, OfficeScanPerspectiveMap mapping) { Image = image; Mapping = mapping; }
    /// <summary>Corrected image; the original source remains unchanged.</summary>
    public OfficeRasterImage Image { get; }
    /// <summary>Projective mapping for overlays and OCR geometry.</summary>
    public OfficeScanPerspectiveMap Mapping { get; }
}

public static partial class OfficeScanProcessor {
    /// <summary>Rectifies an explicitly selected convex page quadrilateral before optional affine and tonal cleanup.</summary>
    /// <remarks>Dimensions follow the longer opposing source edges. This operation does not detect page corners automatically.</remarks>
    public static OfficeScanPerspectiveResult CorrectPerspective(OfficeRasterImage source, OfficeScanPerspectiveOptions options,
        CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (options == null) throw new ArgumentNullException(nameof(options));
        var settings = options.Clone(); settings.Validate(); cancellationToken.ThrowIfCancellationRequested();
        int width = Math.Max(1, (int)Math.Ceiling(Math.Max(Length(settings.TopLeft, settings.TopRight), Length(settings.BottomLeft, settings.BottomRight))));
        int height = Math.Max(1, (int)Math.Ceiling(Math.Max(Length(settings.TopLeft, settings.BottomLeft), Length(settings.TopRight, settings.BottomRight))));
        long sourcePixels = (long)source.Width * source.Height, outputPixels = (long)width * height;
        if (sourcePixels > settings.MaximumPixels || outputPixels > settings.MaximumPixels)
            throw new OfficeScanProcessingLimitException("Perspective correction exceeds MaximumPixels.");
        if ((sourcePixels + outputPixels) * 4L > settings.MaximumWorkingBytes)
            throw new OfficeScanProcessingLimitException("Perspective correction exceeds MaximumWorkingBytes.");
        var mapping = new OfficeScanPerspectiveMap(settings, source.Width, source.Height, width, height);
        return new OfficeScanPerspectiveResult(OfficeRasterResampler.TransformMapped(source, mapping.MapProcessedToSource,
            width, height, OfficeColor.White, cancellationToken), mapping);
        double Length(OfficePoint a, OfficePoint b) {
            double x = (b.X - a.X) * source.Width, y = (b.Y - a.Y) * source.Height;
            return Math.Sqrt(x * x + y * y);
        }
    }
}