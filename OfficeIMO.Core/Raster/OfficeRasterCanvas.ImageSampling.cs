using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    // PDF rasterizers stop filtering uninterpolated upscales at 400%, where pixel replication is intended.
    private const double UninterpolatedUpscaleLimit = 4D;

    /// <summary>
    /// PDF image sampling: box-filter downscaling, bilinear upscaling below 400% (or always with /Interpolate),
    /// pixel replication above it, and axis-aligned placements snapped to whole device pixels.
    /// </summary>
    internal bool PdfImageSampling { get; set; }

    private void DrawPdfImage(OfficeRasterImage image, OfficeTransform imageTransform, OfficeTransform inverse,
        double sourceLeft, double sourceTop, double sourceWidth, double sourceHeight, bool interpolate) {
        (double minX, double minY, double maxX, double maxY) = imageTransform.TransformRectangleBounds(0D, 0D, 1D, 1D);
        if (!IntersectsVisibleImageBounds((minX, minY, maxX, maxY), true)) return;
        image = PrepareImageSource(image);
        double footprintX = SamplingAxisLength(inverse.M11, inverse.M21) * image.Width * sourceWidth;
        double footprintY = SamplingAxisLength(inverse.M12, inverse.M22) * image.Height * sourceHeight;
        image = PrefilterImage(image, footprintX, footprintY);
        bool bilinear = interpolate ||
            footprintX * UninterpolatedUpscaleLimit > 1D && footprintY * UninterpolatedUpscaleLimit > 1D;
        double scaleTolerance = 1e-9D * Math.Max(Math.Abs(imageTransform.M11) + Math.Abs(imageTransform.M22), 1D);
        bool axisAligned = Math.Abs(imageTransform.M12) <= scaleTolerance && Math.Abs(imageTransform.M21) <= scaleTolerance;
        _cancellationToken.ThrowIfCancellationRequested();
        if (axisAligned) {
            SnapSpan(minX, maxX, out double left, out double right);
            SnapSpan(minY, maxY, out double top, out double bottom);
            int x0 = (int)Math.Max(0D, left), x1 = (int)Math.Min(Width, right);
            int y0 = (int)Math.Max(0D, top), y1 = (int)Math.Min(Height, bottom);
            _transformedTextBudget.ChargePaintPixels((long)Math.Max(0, x1 - x0) * Math.Max(0, y1 - y0));
            bool flipX = imageTransform.M11 < 0D, flipY = imageTransform.M22 < 0D;
            for (int py = y0; py < y1; py++) {
                _cancellationToken.ThrowIfCancellationRequested();
                double fy = (py + 0.5D - top) / (bottom - top);
                double v = flipY ? 1D - fy : fy;
                for (int px = x0; px < x1; px++) {
                    double fx = (px + 0.5D - left) / (right - left);
                    double u = flipX ? 1D - fx : fx;
                    BlendPixel(px, py, SamplePdfImage(image, sourceLeft + u * sourceWidth, sourceTop + v * sourceHeight, bilinear));
                }
            }
            return;
        }
        int pixelLeft = Clamp((int)Math.Floor(minX), 0, Width - 1);
        int pixelTop = Clamp((int)Math.Floor(minY), 0, Height - 1);
        int pixelRight = Clamp((int)Math.Ceiling(maxX), 0, Width - 1);
        int pixelBottom = Clamp((int)Math.Ceiling(maxY), 0, Height - 1);
        _transformedTextBudget.ChargePaintPixels((long)(pixelRight - pixelLeft + 1) * (pixelBottom - pixelTop + 1));
        var boundary = new ImageBoundaryCoverage(inverse, 1D, 1D);
        for (int py = pixelTop; py <= pixelBottom; py++) {
            _cancellationToken.ThrowIfCancellationRequested();
            for (int px = pixelLeft; px <= pixelRight; px++) {
                OfficePoint unit = inverse.TransformPoint(new OfficePoint(px + 0.5D, py + 0.5D));
                double coverage = boundary.GetCoverage(unit, out unit);
                if (coverage <= 0D) continue;
                OfficeColor color = SamplePdfImage(image, sourceLeft + unit.X * sourceWidth, sourceTop + unit.Y * sourceHeight, bilinear);
                BlendPixel(px, py, ApplyCoverage(color, coverage));
            }
        }
    }

    private static OfficeColor SamplePdfImage(OfficeRasterImage image, double u, double v, bool bilinear) {
        double sourceX = u * image.Width, sourceY = v * image.Height;
        return bilinear
            ? SampleBilinear(image, sourceX - 0.5D, sourceY - 0.5D)
            : image.GetPixel(Clamp((int)Math.Floor(sourceX), 0, image.Width - 1), Clamp((int)Math.Floor(sourceY), 0, image.Height - 1));
    }

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
        if (ImageCache != null && ImageCache.TryGetFiltered(image, width, height, out OfficeRasterImage cached)) {
            _transformedTextBudget.ChargeIntermediateSurfacePixels((long)cached.Width * cached.Height);
            return cached;
        }
        _transformedTextBudget.ChargeIntermediateSurfacePixels(
            OfficeRasterResampler.GetAdditionalHighQualityPixelBufferCost(image.Width, image.Height, width, height));
        OfficeRasterImage filtered = OfficeRasterResampler.Resize(image, width, height, OfficeRasterResamplingMode.Area,
            OfficeRasterResamplingColorSpace.EncodedSrgb, checked((long)Width * Height * 4L), _cancellationToken);
        ImageCache?.AddFiltered(image, width, height, filtered);
        return filtered;
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
