using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal enum PdfImagePlacementImportDisposition {
    Import,
    SuppressInvisible,
    SuppressUnplaced,
    OmitClippedPixels,
    OmitSoftMask,
    OmitUnsupportedBlendMode,
    OmitUnsupportedPaintEffect,
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
            PdfImagePlacementImportDisposition.SuppressUnplaced;

    internal bool HasNonDefaultOpacity => Opacity < 1D;

    internal bool HasNonNormalBlendMode => BlendMode != OfficeBlendMode.Normal;
}

/// <summary>
/// Keeps reverse-conversion adapters from embedding raw pixels that the source PDF
/// made invisible through placement clipping or an unsupported graphics state.
/// </summary>
internal static class PdfImagePlacementImportPolicy {
    private const double GeometryTolerance = 0.001D;

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
        if (image.SourceImage.HasUnresolvedTransparencyMask) {
            return new PdfImagePlacementImportAssessment(
                PdfImagePlacementImportDisposition.OmitUnresolvedTransparencyMask,
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
        if (!clip.IsExact || !clip.IsRectangle || clip.Width <= 0D || clip.Height <= 0D) return false;

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

    private static double NormalizeOpacity(double opacity) {
        if (double.IsNaN(opacity) || double.IsNegativeInfinity(opacity)) return 0D;
        if (double.IsPositiveInfinity(opacity)) return 1D;
        return Math.Max(0D, Math.Min(1D, opacity));
    }
}
