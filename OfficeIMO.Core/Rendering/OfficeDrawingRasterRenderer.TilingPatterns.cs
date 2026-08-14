namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingRasterRenderer {
    private static void RenderTilingPattern(
        OfficeRasterCanvas canvas,
        OfficeDrawingTilingPattern pattern,
        double scale,
        IOfficeRasterImageCodec? imageCodec,
        long maximumRasterPixels,
        System.Threading.CancellationToken cancellationToken) {
        if (pattern.Opacity <= 0D) return;
        cancellationToken.ThrowIfCancellationRequested();
        _ = OfficeRasterExportPlanner.Resolve(
            pattern.InnerTile.Width,
            pattern.InnerTile.Height,
            OfficeImageExportFormat.Png,
            new OfficeImageExportOptions {
                Scale = scale,
                MaximumRasterPixels = maximumRasterPixels,
                RasterOverflowBehavior = OfficeRasterOverflowBehavior.Throw
            });
        OfficeRasterImage tile = Render(pattern.InnerTile, new OfficeDrawingRasterRenderOptions {
            Scale = scale,
            ImageCodec = imageCodec,
            TextShapingProvider = canvas.TextShapingProvider,
            TextShapingLanguage = canvas.TextShapingLanguage,
            DiagnosticSink = canvas.DiagnosticSink,
            DiagnosticSource = canvas.DiagnosticSource,
            MaximumRasterPixels = maximumRasterPixels,
            CancellationToken = cancellationToken
        });
        bool interpolate = !ContainsNonInterpolatedImage(
            pattern.InnerTile,
            (0D, 0D, pattern.InnerTile.Width, pattern.InnerTile.Height));
        OfficeImagePlacement area = pattern.Area;
        using (canvas.PushClipRectangle(area.X * scale, area.Y * scale, area.Width * scale, area.Height * scale)) {
            foreach (OfficeTransform transform in pattern.GetTileTransforms(pattern.MaximumTileCount)) {
                cancellationToken.ThrowIfCancellationRequested();
                var pixelTransform = new OfficeTransform(transform.M11, transform.M12, transform.M21, transform.M22, transform.OffsetX * scale, transform.OffsetY * scale);
                canvas.DrawAffineImage(tile, pixelTransform, pattern.Opacity, interpolate);
            }
        }
    }
}
