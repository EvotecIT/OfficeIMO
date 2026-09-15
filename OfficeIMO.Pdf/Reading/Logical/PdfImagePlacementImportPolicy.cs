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
    OmitUnappliedDecode,
    OmitUnresolvedTransparencyMask
}

internal readonly struct PdfImagePlacementImportAssessment {
    internal PdfImagePlacementImportAssessment(
        PdfImagePlacementImportDisposition disposition,
        double opacity,
        OfficeBlendMode blendMode) {
        Disposition = disposition;
        Opacity = opacity;
        BlendMode = blendMode;
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
            string.Equals(image.SourceImage.MimeType, "image/jp2", StringComparison.OrdinalIgnoreCase)) {
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
        if (!IsFullyVisibleRectangle(page, placement)) {
            return new PdfImagePlacementImportAssessment(
                PdfImagePlacementImportDisposition.OmitClippedPixels,
                opacity,
                blendMode);
        }

        return new PdfImagePlacementImportAssessment(
            PdfImagePlacementImportDisposition.Import,
            opacity,
            blendMode);
    }

    private static bool IsFullyVisibleRectangle(PdfLogicalPage page, PdfImagePlacement placement) {
        if (placement.Width <= 0D || placement.Height <= 0D) return false;

        PdfSelectionQuad visual = page.MapUserSpaceRectangleToVisual(
            placement.X,
            placement.Y,
            placement.X + placement.Width,
            placement.Y + placement.Height);
        (double pageWidth, double pageHeight) = page.GetVisualPageSize();
        if (visual.Left < -GeometryTolerance ||
            visual.Top < -GeometryTolerance ||
            visual.Right > pageWidth + GeometryTolerance ||
            visual.Bottom > pageHeight + GeometryTolerance) {
            return false;
        }

        PdfImageClipInfo? clip = placement.Clip;
        if (clip == null) return true;
        if (!clip.IsExact ||
            !clip.IsRectangle ||
            clip.ContainsTextClipping ||
            clip.Width <= 0D ||
            clip.Height <= 0D) return false;

        PdfSelectionQuad visualClip = page.MapUserSpaceRectangleToVisual(
            clip.X,
            page.Height - (clip.Y + clip.Height),
            clip.X + clip.Width,
            page.Height - clip.Y);
        return visualClip.Left <= visual.Left + GeometryTolerance &&
            visualClip.Top <= visual.Top + GeometryTolerance &&
            visualClip.Right >= visual.Right - GeometryTolerance &&
            visualClip.Bottom >= visual.Bottom - GeometryTolerance;
    }

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
