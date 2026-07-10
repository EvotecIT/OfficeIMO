namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingRasterRenderer {
    private static void RenderImagePattern(OfficeRasterCanvas canvas, OfficeDrawingImagePattern pattern, double scale) {
        if (!OfficeRasterImageDecoder.TryDecode(pattern.EncodedBytes, out OfficeRasterImage? image) || image == null) {
            return;
        }

        if (pattern.Opacity < 1D) {
            image = ApplyImageOpacity(image, pattern.Opacity);
        }

        OfficeImagePatternLayout layout = pattern.Layout.Scale(scale);
        OfficeImagePlacement area = layout.Area;
        using (canvas.PushClipRectangle(area.X, area.Y, area.Width, area.Height)) {
            foreach (OfficeImagePlacement tile in layout.GetTilePlacements(pattern.MaximumTileCount)) {
                canvas.DrawImage(image, new OfficeImageProjection(tile));
            }
        }
    }
}
