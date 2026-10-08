using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    // A mutable canvas can also be supplied as its own source. Sampling must see
    // the pixels from before drawing, including for overlapping placements.
    private OfficeRasterImage PrepareImageSource(OfficeRasterImage image) {
        if (!ReferenceEquals(_image, image)) return image;
        _cancellationToken.ThrowIfCancellationRequested();
        long pixels = (long)image.Width * image.Height;
        _transformedTextBudget.EnsureIntermediateSurfacePixels(pixels, long.MaxValue);
        OfficeRasterImage snapshot = image.Clone();
        _transformedTextBudget.ChargeIntermediateSurfacePixels(pixels);
        return snapshot;
    }
}
