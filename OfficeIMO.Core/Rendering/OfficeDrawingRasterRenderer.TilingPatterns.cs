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
        canvas = canvas.WithDrawingTextProfile(pattern.InnerTile);
        cancellationToken.ThrowIfCancellationRequested();
        double scaleX = scale * canvas.CoordinateScaleX, scaleY = scale * canvas.CoordinateScaleY;
        _ = OfficeRasterExportPlanner.Resolve(
            pattern.InnerTile.Width * scaleX,
            pattern.InnerTile.Height * scaleY,
            OfficeImageExportFormat.Png,
            new OfficeImageExportOptions {
                Scale = 1D,
                MaximumRasterPixels = maximumRasterPixels,
                RasterOverflowBehavior = OfficeRasterOverflowBehavior.Throw
            });
        canvas.ChargeIntermediateSurfacePixels(
            (long)System.Math.Ceiling(pattern.InnerTile.Width * scaleX) *
            (long)System.Math.Ceiling(pattern.InnerTile.Height * scaleY), maximumRasterPixels);
        OfficeRasterImage tile = RenderCore(pattern.InnerTile, new OfficeDrawingRasterRenderOptions {
            Scale = System.Math.Max(scaleX, scaleY),
            ImageCodec = imageCodec,
            TextShapingProvider = canvas.TextShapingProvider,
            TextShapingLanguage = canvas.TextShapingLanguage,
            DiagnosticSink = canvas.DiagnosticSink,
            DiagnosticSource = canvas.DiagnosticSource,
            TransformedTextBudget = canvas.TransformedTextBudget,
            MaximumRasterPixels = maximumRasterPixels,
            CancellationToken = cancellationToken
        }, scaleX, scaleY);
        bool interpolate = !ContainsNonInterpolatedImage(
            pattern.InnerTile,
            (0D, 0D, pattern.InnerTile.Width, pattern.InnerTile.Height),
            new SamplingInspectionContext(cancellationToken));
        OfficeImagePlacement area = pattern.Area;
        using (canvas.PushClipRectangle(area.X * scale, area.Y * scale, area.Width * scale, area.Height * scale)) {
            foreach (OfficeTransform transform in pattern.GetTileTransforms(pattern.MaximumTileCount)) {
                cancellationToken.ThrowIfCancellationRequested();
                OfficeTransform pixelTransform = new OfficeTransform(
                    transform.M11 * scale / scaleX, transform.M12 * scale / scaleX,
                    transform.M21 * scale / scaleY, transform.M22 * scale / scaleY,
                    transform.OffsetX * scale, transform.OffsetY * scale);
                canvas.DrawAffineImage(tile, pixelTransform, pattern.Opacity, interpolate);
            }
        }
    }
}
