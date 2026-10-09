using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    /// <summary>
    /// Filters source texels over the footprint of a destination pixel before interpolation.
    /// Normalized crop coordinates remain valid on the reduced image. Affine callers scale
    /// their source coordinate transform by the returned image dimensions.
    /// </summary>
    private OfficeRasterImage PrefilterImage(OfficeRasterImage image, double footprintX, double footprintY) {
        int width = ReducedSamplingDimension(image.Width, footprintX);
        int height = ReducedSamplingDimension(image.Height, footprintY);
        if (width == image.Width && height == image.Height) return image;
        _cancellationToken.ThrowIfCancellationRequested();
        _transformedTextBudget.ChargeIntermediateSurfacePixels(
            OfficeRasterResampler.GetAdditionalHighQualityPixelBufferCost(image.Width, image.Height, width, height));
        return OfficeRasterResampler.Resize(image, width, height, OfficeRasterResamplingMode.Area,
            OfficeRasterResamplingColorSpace.EncodedSrgb, checked((long)Width * Height * 4L), _cancellationToken);
    }

    private static int ReducedSamplingDimension(int original, double footprint) =>
        !IsFinite(footprint) || footprint <= 1D ? original : Math.Max(1, (int)Math.Ceiling(original / footprint));

    private static double SamplingAxisLength(double x, double y) {
        double largest = Math.Max(Math.Abs(x), Math.Abs(y));
        if (largest == 0D) return 0D;
        return largest * Math.Sqrt((x / largest) * (x / largest) + (y / largest) * (y / largest));
    }

    // Actual source-space filtered texel dimensions used by DrawAffineImage.
    internal static (double X, double Y) AffineImageSamplingStep(int width, int height, OfficeTransform transform) {
        if (!transform.TryInvert(out OfficeTransform inverse)) return (1D, 1D);
        int filteredWidth = ReducedSamplingDimension(width, SamplingAxisLength(inverse.M11, inverse.M21));
        int filteredHeight = ReducedSamplingDimension(height, SamplingAxisLength(inverse.M12, inverse.M22));
        return (width / (double)filteredWidth, height / (double)filteredHeight);
    }

    private OfficeRasterImage PrefilterAffineImage(OfficeRasterImage image, ref OfficeTransform inverse) {
        OfficeRasterImage filtered = PrefilterImage(image, SamplingAxisLength(inverse.M11, inverse.M21), SamplingAxisLength(inverse.M12, inverse.M22));
        if (!ReferenceEquals(filtered, image)) inverse = inverse.Then(OfficeTransform.Scale(filtered.Width / (double)image.Width, filtered.Height / (double)image.Height));
        return filtered;
    }
}
