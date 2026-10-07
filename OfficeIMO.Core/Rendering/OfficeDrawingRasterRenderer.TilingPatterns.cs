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
        OfficeImagePlacement area = pattern.Area;
        using var clip = canvas.PushClipRectangle(area.X * scale, area.Y * scale, area.Width * scale, area.Height * scale);
        if (!canvas.HasVisibleClipBounds) return;
        var transforms = pattern.GetTileTransforms(pattern.MaximumTileCount);
        if (transforms.Count == 0) return;
        bool interpolate = !ContainsNonInterpolatedImage(
            pattern.InnerTile,
            (0D, 0D, pattern.InnerTile.Width, pattern.InnerTile.Height),
            new SamplingInspectionContext(cancellationToken));
        // Tile axes are transformed into the parent canvas before choosing their
        // density. Nearest content keeps its original grid, as effect layers do.
        (double axisX, double axisY) = interpolate
            ? GetEffectAxisScales(pattern.Transform, canvas.CoordinateScaleX, canvas.CoordinateScaleY)
            : (1D, 1D);
        double scaleX = scale * axisX, scaleY = scale * axisY;
        double width = System.Math.Ceiling(pattern.InnerTile.Width * scaleX);
        double height = System.Math.Ceiling(pattern.InnerTile.Height * scaleY);
        bool visibleTile = double.IsInfinity(width) || double.IsInfinity(height);
        foreach (OfficeTransform transform in transforms) {
            cancellationToken.ThrowIfCancellationRequested();
            if (visibleTile || canvas.IntersectsVisibleSurface(CreateTilePixelTransform(transform, scale, scaleX, scaleY), width, height)) {
                visibleTile = true;
                break;
            }
        }
        if (!visibleTile) return;
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
        foreach (OfficeTransform transform in transforms) {
            cancellationToken.ThrowIfCancellationRequested();
            canvas.DrawAffineImage(tile, CreateTilePixelTransform(transform, scale, scaleX, scaleY), pattern.Opacity, interpolate);
        }
    }

    private static OfficeTransform CreateTilePixelTransform(OfficeTransform transform, double scale, double scaleX, double scaleY) =>
        new OfficeTransform(transform.M11 * scale / scaleX, transform.M12 * scale / scaleX,
            transform.M21 * scale / scaleY, transform.M22 * scale / scaleY,
            transform.OffsetX * scale, transform.OffsetY * scale);
}
