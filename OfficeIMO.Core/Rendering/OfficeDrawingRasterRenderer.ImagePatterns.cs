namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingRasterRenderer {
    private static void RenderImagePattern(
        OfficeRasterCanvas canvas,
        OfficeDrawingImagePattern pattern,
        double scale,
        IOfficeRasterImageCodec? imageCodec,
        long maximumRasterPixels,
        System.Threading.CancellationToken cancellationToken) {
        if (pattern.Opacity <= 0D) return;
        cancellationToken.ThrowIfCancellationRequested();
        OfficeImagePatternLayout layout = pattern.Layout.Scale(scale);
        OfficeImagePlacement area = layout.Area;
        using var clip = canvas.PushClipRectangle(area.X, area.Y, area.Width, area.Height);
        if (!canvas.HasVisibleClipBounds) return;
        var placements = layout.GetTilePlacements(pattern.MaximumTileCount);
        bool visibleTile = false;
        foreach (OfficeImagePlacement placement in placements) {
            cancellationToken.ThrowIfCancellationRequested();
            if (canvas.IntersectsVisibleSurface(new OfficeImageProjection(placement).CreateUnitSquareTransform(), 1D, 1D)) {
                visibleTile = true;
                break;
            }
        }
        if (!visibleTile) return;
        (double targetWidth, double targetHeight) = GetImageTargetSize(canvas, new OfficeImageProjection(layout.Tile), 1D);
        if (!TryDecodeImage(
                pattern.EncodedBytes,
                pattern.ContentType,
                targetWidth,
                targetHeight,
                imageCodec,
                canvas.TextShapingProvider,
                canvas.TextShapingLanguage,
                canvas.DiagnosticSink,
                canvas.DiagnosticSource,
                canvas.TransformedTextBudget,
                maximumRasterPixels,
                cancellationToken,
                out OfficeRasterImage? image) ||
            image == null) {
            return;
        }

        if (pattern.Opacity < 1D) {
            canvas.ChargeIntermediateSurfacePixels((long)image.Width * image.Height, maximumRasterPixels);
            image = ApplyImageOpacity(image, pattern.Opacity);
        }

        foreach (OfficeImagePlacement tile in placements) {
            cancellationToken.ThrowIfCancellationRequested();
            canvas.DrawImage(image, new OfficeImageProjection(tile));
        }
    }
}
