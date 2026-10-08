using System;

namespace OfficeIMO.Drawing;

// Shared by nested canvases and image decoders. Remember the operation's limit so
// sampling buffers admitted by a canvas use the same ceiling as decoded images.
internal sealed class OfficeRasterTransformedTextBudget {
    internal long Pixels;
    internal long IntermediatePixels;
    private long _maximumRasterPixels = long.MaxValue;

    internal long GetRemainingIntermediateSurfacePixels(long maximumRasterPixels) =>
        Math.Min(_maximumRasterPixels, maximumRasterPixels) - IntermediatePixels;

    internal void EnsureIntermediateSurfacePixels(long pixels, long maximumRasterPixels) {
        maximumRasterPixels = Math.Min(_maximumRasterPixels, maximumRasterPixels);
        long consumed = IntermediatePixels;
        if (pixels < 0L || pixels > GetRemainingIntermediateSurfacePixels(maximumRasterPixels)) {
            throw new OfficeImageExportLimitException(1D,
                pixels > long.MaxValue - consumed ? long.MaxValue : consumed + pixels,
                maximumRasterPixels,
                OfficeRasterImageEncoder.GetMaximumDimension(OfficeImageExportFormat.Png));
        }
    }

    internal void ChargeIntermediateSurfacePixels(long pixels) =>
        ChargeIntermediateSurfacePixels(pixels, _maximumRasterPixels);

    internal void ChargeIntermediateSurfacePixels(long pixels, long maximumRasterPixels) {
        EnsureIntermediateSurfacePixels(pixels, maximumRasterPixels);
        _maximumRasterPixels = Math.Min(_maximumRasterPixels, maximumRasterPixels);
        IntermediatePixels += pixels;
    }

    internal void ReleaseIntermediateSurfacePixels(long pixels) => IntermediatePixels -= pixels;
}
