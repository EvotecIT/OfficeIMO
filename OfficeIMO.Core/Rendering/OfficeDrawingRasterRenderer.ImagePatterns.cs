namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingRasterRenderer {
    private static void RenderImagePattern(
        OfficeRasterCanvas canvas,
        OfficeDrawingImagePattern pattern,
        double scale,
        IOfficeRasterImageCodec? imageCodec,
        System.Threading.CancellationToken cancellationToken) {
        OfficeImagePatternLayout layout = pattern.Layout.Scale(scale);
        if (!TryDecodeImage(
                pattern.EncodedBytes,
                pattern.ContentType,
                layout.Tile.Width,
                layout.Tile.Height,
                imageCodec,
                canvas.TextShapingProvider,
                canvas.TextShapingLanguage,
                canvas.DiagnosticSink,
                canvas.DiagnosticSource,
                cancellationToken,
                out OfficeRasterImage? image) ||
            image == null) {
            return;
        }

        if (pattern.Opacity < 1D) {
            image = ApplyImageOpacity(image, pattern.Opacity);
        }

        OfficeImagePlacement area = layout.Area;
        using (canvas.PushClipRectangle(area.X, area.Y, area.Width, area.Height)) {
            foreach (OfficeImagePlacement tile in layout.GetTilePlacements(pattern.MaximumTileCount)) {
                canvas.DrawImage(image, new OfficeImageProjection(tile));
            }
        }
    }
}
