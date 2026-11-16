using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal enum PdfImagePlacementImportDisposition {
    Import,
    SuppressInvisible,
    SuppressOutsideVisibleArea,
    SuppressUnplaced,
    OmitClippedPixels,
    OmitSoftMask,
    OmitUnsupportedBlendMode,
    OmitUnsupportedPaintEffect,
    OmitUnsupportedTransform,
    OmitUnappliedDecode,
    OmitUnresolvedTransparencyMask
}

internal readonly struct PdfImagePlacementImportAssessment {
    internal PdfImagePlacementImportAssessment(
        PdfImagePlacementImportDisposition disposition,
        double opacity,
        OfficeBlendMode blendMode,
        PdfImagePlacementCrop? crop = null) {
        Disposition = disposition;
        Opacity = opacity;
        BlendMode = blendMode;
        Crop = crop;
    }

    /// <summary>
    /// The rectangular clip that cuts into the placed image, or null when the whole image is visible.
    /// Adapters embed <see cref="PdfImagePlacementImportPolicy.TryGetVisibleImagePayload"/> and place the
    /// visible bounds, so pixels hidden by the clip never reach the output.
    /// </summary>
    internal PdfImagePlacementCrop? Crop { get; }

    /// <summary>Visible placement bounds in page user space (bottom-left origin): the crop when present, otherwise the placement.</summary>
    internal (double X, double Y, double Width, double Height) GetVisibleBounds(PdfImagePlacement placement) {
        Guard.NotNull(placement, nameof(placement));
        return Crop is { } crop
            ? (crop.X, crop.Y, crop.Width, crop.Height)
            : (placement.X, placement.Y, placement.Width, placement.Height);
    }

    internal PdfImagePlacementImportDisposition Disposition { get; }

    internal double Opacity { get; }

    internal OfficeBlendMode BlendMode { get; }

    internal bool CanImport => Disposition == PdfImagePlacementImportDisposition.Import;

    internal bool IsSuppressed =>
        Disposition is PdfImagePlacementImportDisposition.SuppressInvisible or
            PdfImagePlacementImportDisposition.SuppressOutsideVisibleArea or
            PdfImagePlacementImportDisposition.SuppressUnplaced;

    internal bool HasNonDefaultOpacity => Opacity < 1D;

    internal int MappedTransparencyPercent {
        get {
            int percentage = (int)Math.Round((1D - Opacity) * 100D, MidpointRounding.AwayFromZero);
            if (percentage < 0) return 0;
            return percentage > 100 ? 100 : percentage;
        }
    }

    internal bool MappedOpacityIsOmitted => Opacity > 0D && MappedTransparencyPercent >= 100;

    internal bool HasNonNormalBlendMode => BlendMode != OfficeBlendMode.Normal;
}

/// <summary>
/// The visible part of an image placement whose rectangular clip cuts into the image. Bounds are in page user space
/// with a bottom-left origin, like <see cref="PdfImagePlacement"/>; source fractions are measured on the stored image.
/// </summary>
internal sealed class PdfImagePlacementCrop {
    internal PdfImagePlacementCrop(
        double x,
        double y,
        double width,
        double height,
        double sourceLeft,
        double sourceTop,
        double sourceRight,
        double sourceBottom) {
        X = x;
        Y = y;
        Width = width;
        Height = height;
        SourceLeft = sourceLeft;
        SourceTop = sourceTop;
        SourceRight = sourceRight;
        SourceBottom = sourceBottom;
    }

    internal double X { get; }

    internal double Y { get; }

    internal double Width { get; }

    internal double Height { get; }

    internal double SourceLeft { get; }

    internal double SourceTop { get; }

    internal double SourceRight { get; }

    internal double SourceBottom { get; }
}

/// <summary>
/// Keeps reverse-conversion adapters from embedding raw pixels that the source PDF
/// made invisible through placement clipping or an unsupported graphics state.
/// </summary>
internal static class PdfImagePlacementImportPolicy {
    private const double GeometryTolerance = 0.001D;

    internal static bool HasVisiblePlacement(PdfLogicalPage page, PdfLogicalImage image) {
        Guard.NotNull(page, nameof(page));
        Guard.NotNull(image, nameof(image));

        for (int placementIndex = 0; placementIndex < image.Placements.Count; placementIndex++) {
            if (!Analyze(page, image, image.Placements[placementIndex]).IsSuppressed) return true;
        }
        return false;
    }

    internal static PdfImagePlacementImportAssessment Analyze(
        PdfLogicalPage page,
        PdfLogicalImage image,
        PdfImagePlacement? placement) {
        Guard.NotNull(page, nameof(page));
        Guard.NotNull(image, nameof(image));

        if (placement == null) {
            return new PdfImagePlacementImportAssessment(
                PdfImagePlacementImportDisposition.SuppressUnplaced,
                1D,
                OfficeBlendMode.Normal);
        }

        double opacity = NormalizeOpacity(placement.Opacity);
        OfficeBlendMode blendMode = placement.EffectiveBlendMode;
        if (opacity <= 0D) {
            return new PdfImagePlacementImportAssessment(
                PdfImagePlacementImportDisposition.SuppressInvisible,
                opacity,
                blendMode);
        }
        if (!HasVisibleIntersection(page, placement)) {
            return new PdfImagePlacementImportAssessment(
                PdfImagePlacementImportDisposition.SuppressOutsideVisibleArea,
                opacity,
                blendMode);
        }
        if (image.SourceImage.HasUnresolvedTransparencyMask) {
            return new PdfImagePlacementImportAssessment(
                PdfImagePlacementImportDisposition.OmitUnresolvedTransparencyMask,
                opacity,
                blendMode);
        }
        if (image.SourceImage.HasUnsafePassThroughDecode &&
            IsPassThroughScanMimeType(image.SourceImage.MimeType)) {
            return new PdfImagePlacementImportAssessment(
                PdfImagePlacementImportDisposition.OmitUnappliedDecode,
                opacity,
                blendMode);
        }
        if (placement.HasSoftMask) {
            return new PdfImagePlacementImportAssessment(
                PdfImagePlacementImportDisposition.OmitSoftMask,
                opacity,
                blendMode);
        }
        if (placement.HasUnsupportedBlendMode) {
            return new PdfImagePlacementImportAssessment(
                PdfImagePlacementImportDisposition.OmitUnsupportedBlendMode,
                opacity,
                blendMode);
        }
        if (placement.HasUnsupportedImagePaintEffect) {
            return new PdfImagePlacementImportAssessment(
                PdfImagePlacementImportDisposition.OmitUnsupportedPaintEffect,
                opacity,
                blendMode);
        }
        if (!placement.IsAxisAligned || !IsFinite(placement.B) || !IsFinite(placement.C)) {
            return new PdfImagePlacementImportAssessment(
                PdfImagePlacementImportDisposition.OmitUnsupportedTransform,
                opacity,
                blendMode);
        }
        if (!TryResolveVisibleRectangle(page, image, placement, out PdfImagePlacementCrop? crop)) {
            return new PdfImagePlacementImportAssessment(
                PdfImagePlacementImportDisposition.OmitClippedPixels,
                opacity,
                blendMode);
        }

        return new PdfImagePlacementImportAssessment(
            PdfImagePlacementImportDisposition.Import,
            opacity,
            blendMode,
            crop);
    }

    /// <summary>
    /// Returns the image file an adapter embeds for an importable placement. Without a crop this is the extracted
    /// payload unchanged. With a crop the pixels are decoded, cut to the clip and written as PNG, so the output never
    /// carries pixels the PDF hides. False means the payload could not be cropped and the image must be omitted.
    /// </summary>
    internal static bool TryGetVisibleImagePayload(
        PdfLogicalImage image,
        PdfImagePlacementImportAssessment assessment,
        out byte[] bytes,
        out string mimeType,
        out string fileExtension) {
        Guard.NotNull(image, nameof(image));
        PdfExtractedImage source = image.SourceImage;
        PdfImagePlacementCrop? crop = assessment.Crop;
        if (crop == null) {
            bytes = source.Bytes;
            mimeType = source.MimeType ?? string.Empty;
            fileExtension = source.FileExtension ?? string.Empty;
            return bytes.Length > 0;
        }

        bytes = Array.Empty<byte>();
        mimeType = "image/png";
        fileExtension = ".png";
        if (!source.IsImageFile) return false;
        if (!OfficeRasterImageDecoder.TryDecode(source.EncodedBytes, out OfficeRasterImage? raster) || raster == null) return false;

        int left = ToPixels(crop.SourceLeft, raster.Width);
        int right = ToPixels(crop.SourceRight, raster.Width);
        int top = ToPixels(crop.SourceTop, raster.Height);
        int bottom = ToPixels(crop.SourceBottom, raster.Height);
        int width = raster.Width - left - right;
        int height = raster.Height - top - bottom;
        if (width <= 0 || height <= 0) return false;

        OfficeRasterImage visible = OfficeRasterTransforms.Crop(raster, left, top, width, height);
        bytes = OfficePngWriter.Encode(visible);
        return bytes.Length > 0;
    }

    /// <summary>
    /// Uses the source alternative text or the detected figure caption. Internal resource names describe nothing to
    /// a reader, so an image without either gets an empty description rather than a synthetic one.
    /// </summary>
    internal static string GetAlternativeText(PdfLogicalImage image, PdfImagePlacement? placement) {
        Guard.NotNull(image, nameof(image));
        PdfUnderstandingImageRegion? region = null;
        for (int index = 0; index < image.Regions.Count; index++) {
            if (placement == null || ReferenceEquals(image.Regions[index].Placement, placement)) {
                region = image.Regions[index];
                break;
            }
        }
        region ??= image.PrimaryRegion;
        if (!string.IsNullOrWhiteSpace(region?.AlternativeText)) return region!.AlternativeText!.Trim();
        string? caption = region?.Caption?.Region.Text;
        return string.IsNullOrWhiteSpace(caption) ? string.Empty : caption!.Trim();
    }

    /// <summary>Pixel size of the payload <see cref="TryGetVisibleImagePayload"/> embeds for an image of the given size.</summary>
    internal static (int Width, int Height) GetVisiblePixelSize(int width, int height, PdfImagePlacementImportAssessment assessment) {
        if (assessment.Crop is not { } crop) return (width, height);
        return (
            Math.Max(1, width - ToPixels(crop.SourceLeft, width) - ToPixels(crop.SourceRight, width)),
            Math.Max(1, height - ToPixels(crop.SourceTop, height) - ToPixels(crop.SourceBottom, height)));
    }

    private static int ToPixels(double fraction, int pixels) =>
        Math.Max(0, Math.Min(pixels, (int)Math.Round(fraction * pixels, MidpointRounding.AwayFromZero)));

    private static bool IsPassThroughScanMimeType(string? mimeType) =>
        string.Equals(mimeType, "image/jpeg", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(mimeType, "image/jp2", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(mimeType, "image/j2c", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Resolves the part of an axis-aligned placement that the page shows. A rectangular clip that contains the
    /// image, or trims it by less than half a source pixel (office suites commonly clip every image to its own
    /// rounded frame), leaves the image whole. A rectangular clip that cuts deeper becomes a crop. Clips that
    /// are not exact rectangles, text clips, and images still partly off the page after clipping are refused.
    /// </summary>
    private static bool TryResolveVisibleRectangle(
        PdfLogicalPage page,
        PdfLogicalImage image,
        PdfImagePlacement placement,
        out PdfImagePlacementCrop? crop) {
        crop = null;
        if (placement.Width <= 0D || placement.Height <= 0D) return false;

        double left = placement.X;
        double bottom = placement.Y;
        double right = placement.X + placement.Width;
        double top = placement.Y + placement.Height;

        PdfImageClipInfo? clip = placement.Clip;
        if (clip != null) {
            if (!clip.IsExact ||
                !clip.IsRectangle ||
                clip.ContainsTextClipping ||
                !IsFinite(clip.X) || !IsFinite(clip.Y) ||
                !IsFinite(clip.Width) || !IsFinite(clip.Height) ||
                clip.Width <= 0D ||
                clip.Height <= 0D) return false;

            double clipLeft = clip.X;
            double clipRight = clip.X + clip.Width;
            double clipBottom = page.Height - (clip.Y + clip.Height);
            double clipTop = page.Height - clip.Y;

            // A trim smaller than half a source pixel does not remove a visible pixel.
            double pixelWidth = image.Width > 0 ? placement.Width / image.Width : 0D;
            double pixelHeight = image.Height > 0 ? placement.Height / image.Height : 0D;
            double toleranceX = Math.Max(GeometryTolerance, pixelWidth / 2D);
            double toleranceY = Math.Max(GeometryTolerance, pixelHeight / 2D);
            double trimLeft = Trim(clipLeft - left, toleranceX);
            double trimRight = Trim(right - clipRight, toleranceX);
            double trimBottom = Trim(clipBottom - bottom, toleranceY);
            double trimTop = Trim(top - clipTop, toleranceY);

            if (trimLeft > 0D || trimRight > 0D || trimBottom > 0D || trimTop > 0D) {
                if (trimLeft + trimRight >= placement.Width - toleranceX ||
                    trimBottom + trimTop >= placement.Height - toleranceY) return false;

                left += trimLeft;
                right -= trimRight;
                bottom += trimBottom;
                top -= trimTop;

                // Fractions follow the stored pixel order. Image row 0 is at the top of the unit square, and a
                // negative scale mirrors the image, so a trimmed user-space edge maps to the opposite source edge.
                double fractionLeft = trimLeft / placement.Width;
                double fractionRight = trimRight / placement.Width;
                double fractionTop = trimTop / placement.Height;
                double fractionBottom = trimBottom / placement.Height;
                crop = new PdfImagePlacementCrop(
                    left,
                    bottom,
                    right - left,
                    top - bottom,
                    placement.A < 0D ? fractionRight : fractionLeft,
                    placement.D < 0D ? fractionBottom : fractionTop,
                    placement.A < 0D ? fractionLeft : fractionRight,
                    placement.D < 0D ? fractionTop : fractionBottom);

                // A clip that leaves less than one whole source pixel has nothing safe to show.
                if ((image.Width > 0 && image.Width - ToPixels(crop.SourceLeft, image.Width) - ToPixels(crop.SourceRight, image.Width) < 1) ||
                    (image.Height > 0 && image.Height - ToPixels(crop.SourceTop, image.Height) - ToPixels(crop.SourceBottom, image.Height) < 1)) {
                    crop = null;
                    return false;
                }
            }
        }

        PdfSelectionQuad visual = page.MapUserSpaceRectangleToVisual(left, bottom, right, top);
        (double pageWidth, double pageHeight) = page.GetVisualPageSize();
        return visual.Left >= -GeometryTolerance &&
            visual.Top >= -GeometryTolerance &&
            visual.Right <= pageWidth + GeometryTolerance &&
            visual.Bottom <= pageHeight + GeometryTolerance;
    }

    private static double Trim(double value, double tolerance) => value > tolerance ? value : 0D;

    private static bool HasVisibleIntersection(PdfLogicalPage page, PdfImagePlacement placement) {
        if (!IsFinite(placement.X) || !IsFinite(placement.Y) ||
            !IsFinite(placement.Width) || !IsFinite(placement.Height) ||
            placement.Width <= 0D || placement.Height <= 0D ||
            !IsFinite(placement.X + placement.Width) || !IsFinite(placement.Y + placement.Height)) return false;

        PdfSelectionQuad visual = page.MapUserSpaceRectangleToVisual(
            placement.X,
            placement.Y,
            placement.X + placement.Width,
            placement.Y + placement.Height);
        (double pageWidth, double pageHeight) = page.GetVisualPageSize();
        double left = Math.Max(0D, visual.Left);
        double top = Math.Max(0D, visual.Top);
        double right = Math.Min(pageWidth, visual.Right);
        double bottom = Math.Min(pageHeight, visual.Bottom);
        if (right <= left + GeometryTolerance || bottom <= top + GeometryTolerance) return false;

        PdfImageClipInfo? clip = placement.Clip;
        if (clip == null) return true;
        if (!IsFinite(clip.X) || !IsFinite(clip.Y) ||
            !IsFinite(clip.Width) || !IsFinite(clip.Height) ||
            clip.Width <= 0D || clip.Height <= 0D ||
            !IsFinite(clip.X + clip.Width) || !IsFinite(clip.Y + clip.Height)) return false;

        PdfSelectionQuad visualClip = page.MapUserSpaceRectangleToVisual(
            clip.X,
            page.Height - (clip.Y + clip.Height),
            clip.X + clip.Width,
            page.Height - clip.Y);
        left = Math.Max(left, visualClip.Left);
        top = Math.Max(top, visualClip.Top);
        right = Math.Min(right, visualClip.Right);
        bottom = Math.Min(bottom, visualClip.Bottom);
        return right > left + GeometryTolerance && bottom > top + GeometryTolerance;
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static double NormalizeOpacity(double opacity) {
        if (double.IsNaN(opacity) || double.IsNegativeInfinity(opacity)) return 0D;
        if (double.IsPositiveInfinity(opacity)) return 1D;
        return Math.Max(0D, Math.Min(1D, opacity));
    }
}
