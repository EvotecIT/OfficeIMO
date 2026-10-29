using System;

namespace OfficeIMO.Drawing;

// Shared by nested canvases and image decoders. Remember the operation's limit so
// sampling buffers admitted by a canvas use the same ceiling as decoded images.
internal sealed class OfficeRasterTransformedTextBudget {
    internal long Pixels;
    internal long IntermediatePixels;
    private long _maximumRasterPixels = long.MaxValue;
    private long? _remainingPaintPixels;
    private long _remainingPaintCrossings;

    internal void LimitPaintWork(long outputPixels) {
        // Count actual clipped raster work, preserving the native SVG reader's grammar.
        // Nested effects, patterns, images and text share this operation's budget.
        _remainingPaintPixels = outputPixels > long.MaxValue / 256L ? long.MaxValue : outputPixels * 256L;
        _remainingPaintCrossings = 512_000_000L;
    }

    internal void ChargePaintPixels(long pixels) {
        if (!_remainingPaintPixels.HasValue) return;
        if (pixels < 0L || pixels > _remainingPaintPixels.Value)
            throw new NotSupportedException("Cumulative SVG raster paint area exceeds the image normalization limit.");
        _remainingPaintPixels -= pixels;
    }

    internal void ChargePaintCrossings(long crossings) {
        if (!_remainingPaintPixels.HasValue) return;
        if (crossings < 0L || crossings > _remainingPaintCrossings)
            throw new NotSupportedException("Cumulative SVG contour work exceeds the image normalization limit.");
        _remainingPaintCrossings -= crossings;
    }

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
