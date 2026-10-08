using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private const long MaximumTransformedTextIntermediatePixels = 64_000_000L;
    private OfficeRasterTransformedTextBudget _transformedTextBudget = new OfficeRasterTransformedTextBudget();

    // Optional sampling support must fit both the render-wide allocation ceiling
    // and the cumulative text ceiling before the required layer is charged.
    internal long GetRemainingTransformedTextIntermediatePixels(long maximumRasterPixels) =>
        Math.Min(MaximumTransformedTextIntermediatePixels - _transformedTextBudget.Pixels,
            _transformedTextBudget.GetRemainingIntermediateSurfacePixels(maximumRasterPixels));

    internal void ChargeTransformedTextIntermediatePixels(long pixels, long maximumRasterPixels) {
        long consumed = _transformedTextBudget.Pixels;
        if (pixels < 0L || pixels > MaximumTransformedTextIntermediatePixels - consumed) {
            throw new OfficeImageExportLimitException(1D,
                pixels > long.MaxValue - consumed ? long.MaxValue : consumed + pixels,
                MaximumTransformedTextIntermediatePixels,
                OfficeRasterImageEncoder.GetMaximumDimension(OfficeImageExportFormat.Png));
        }
        _transformedTextBudget.ChargeIntermediateSurfacePixels(pixels, maximumRasterPixels);
        _transformedTextBudget.Pixels = consumed + pixels;
    }

    internal void ReleaseTransformedTextIntermediatePixels(long pixels) {
        _transformedTextBudget.Pixels -= pixels;
        _transformedTextBudget.ReleaseIntermediateSurfacePixels(pixels);
    }

    internal OfficeRasterTransformedTextBudget TransformedTextBudget => _transformedTextBudget;

    internal void ChargeIntermediateSurfacePixels(long pixels, long maximumRasterPixels) =>
        _transformedTextBudget.ChargeIntermediateSurfacePixels(pixels, maximumRasterPixels);

    internal void ShareTransformedTextBudget(OfficeRasterTransformedTextBudget budget) =>
        _transformedTextBudget = budget;
}
