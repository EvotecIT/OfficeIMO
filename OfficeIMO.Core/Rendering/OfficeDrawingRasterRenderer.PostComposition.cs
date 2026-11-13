namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingRasterRenderer {
    // A pointwise destination transform needs only the visible source texels. Retain the
    // original grid and a bilinear halo so nested page tiles do not allocate a whole page.
    private static bool TryBoundPostCompositionLayer(OfficeRasterCanvas canvas, OfficeDrawingEffectGroup effect,
        ref EffectTextLayerPlan layer, ref OfficeTransform pixelTransform, double fullWidth, double fullHeight,
        out int width, out int height) {
        width = 0; height = 0;
        if (layer.Drawing.PostCompositionColorTransform == null || effect.SoftMask != null ||
            fullWidth <= 0D || fullHeight <= 0D || fullWidth > int.MaxValue || fullHeight > int.MaxValue ||
            double.IsNaN(fullWidth) || double.IsNaN(fullHeight)) return false;
        OfficeTransform physical = pixelTransform.Then(OfficeTransform.Scale(canvas.CoordinateScaleX, canvas.CoordinateScaleY));
        if (!physical.TryInvert(out OfficeTransform inverse)) return false;
        // Area prefiltering and external masks use their complete source grid. Preserve
        // that existing isolation and its budget rejection rather than changing sampling.
        if (layer.Interpolate) {
            var step = OfficeRasterCanvas.AffineImageSamplingStep((int)fullWidth, (int)fullHeight, physical);
            if (step.X != 1D || step.Y != 1D) return false;
        }
        var visible = inverse.TransformRectangleBounds(0D, 0D, canvas.Width, canvas.Height);
        if (!IsFiniteEffectAllocation(visible.Left) || !IsFiniteEffectAllocation(visible.Top) ||
            !IsFiniteEffectAllocation(visible.Right) || !IsFiniteEffectAllocation(visible.Bottom)) return false;
        double left = Math.Max(0D, Math.Floor(visible.Left) - 2D);
        double top = Math.Max(0D, Math.Floor(visible.Top) - 2D);
        double right = Math.Min(fullWidth, Math.Ceiling(visible.Right) + 2D);
        double bottom = Math.Min(fullHeight, Math.Ceiling(visible.Bottom) + 2D);
        if (right <= left || bottom <= top ||
            (left == 0D && top == 0D && right == fullWidth && bottom == fullHeight)) return false;
        width = (int)(right - left); height = (int)(bottom - top);
        OfficeDrawing drawing = layer.Drawing;
        var viewport = new OfficeDrawing(width / layer.ScaleX, height / layer.ScaleY) {
            TextShapingProvider = drawing.TextShapingProvider,
            TextShapingLanguage = drawing.TextShapingLanguage
        };
        viewport.AddDrawingForClippedRendering(drawing.CreateCompositionSource(), -left / layer.ScaleX, -top / layer.ScaleY, null);
        viewport.SetPostCompositionColorTransform(drawing.PostCompositionColorTransform!);
        layer = new EffectTextLayerPlan(viewport, 0D, layer.ScaleX, layer.ScaleY, layer.Interpolate, layer.ParagraphInk);
        pixelTransform = OfficeTransform.Translate(left, top).Then(pixelTransform);
        return true;
    }

    private static void ApplyPostCompositionColorTransform(OfficeRasterImage image,
        IOfficePostCompositionColorTransform? transform, CancellationToken cancellationToken) {
        if (transform == null) return;
        byte[] pixels = image.PixelBuffer;
        int stride = checked(image.Width * 4);
        for (int y = 0; y < image.Height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            int rowEnd = checked((y + 1) * stride);
            for (int offset = y * stride; offset < rowEnd; offset += 4) {
                if ((offset & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                byte alpha = pixels[offset + 3];
                if (alpha == 0) continue;
                OfficeColor source = OfficeColor.FromRgba(pixels[offset], pixels[offset + 1], pixels[offset + 2], alpha);
                if (!transform.TryTransform(source, out OfficeColor converted)) continue;
                pixels[offset] = converted.R;
                pixels[offset + 1] = converted.G;
                pixels[offset + 2] = converted.B;
                // A color transform cannot change surface coverage.
            }
        }
    }
}
