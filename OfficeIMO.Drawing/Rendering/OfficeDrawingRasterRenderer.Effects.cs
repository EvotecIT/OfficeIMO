namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingRasterRenderer {
    private static void RenderEffectGroup(
        OfficeRasterCanvas canvas,
        OfficeDrawingEffectGroup effectGroup,
        double scale,
        IOfficeRasterImageCodec? imageCodec,
        System.Threading.CancellationToken cancellationToken) {
        if (effectGroup.Opacity <= 0D) return;
        cancellationToken.ThrowIfCancellationRequested();
        OfficeRasterImage layer = Render(effectGroup.InnerDrawing, new OfficeDrawingRasterRenderOptions {
            Scale = scale,
            ImageCodec = imageCodec,
            CancellationToken = cancellationToken
        });
        if (effectGroup.SoftMask != null) layer = ApplySoftMask(layer, effectGroup.SoftMask, scale, imageCodec, cancellationToken);
        OfficeTransform transform = effectGroup.Transform;
        var pixelTransform = new OfficeTransform(transform.M11, transform.M12, transform.M21, transform.M22, transform.OffsetX * scale, transform.OffsetY * scale);
        canvas.DrawAffineImage(layer, pixelTransform, effectGroup.Opacity, effectGroup.BlendMode);
    }

    private static OfficeRasterImage ApplySoftMask(
        OfficeRasterImage source,
        OfficeDrawingSoftMask softMask,
        double scale,
        IOfficeRasterImageCodec? imageCodec,
        System.Threading.CancellationToken cancellationToken) {
        var maskScene = new OfficeDrawing(source.Width / scale, source.Height / scale);
        maskScene.AddEffectDrawing(softMask.InnerDrawing, softMask.Transform);
        OfficeRasterImage mask = Render(maskScene, new OfficeDrawingRasterRenderOptions {
            Scale = scale,
            ImageCodec = imageCodec,
            CancellationToken = cancellationToken
        });
        var result = new OfficeRasterImage(source.Width, source.Height);
        double backdrop = GetMaskFactor(softMask.BackdropColor, softMask.Mode);
        for (int y = 0; y < source.Height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int x = 0; x < source.Width; x++) {
                OfficeColor sourcePixel = source.GetPixel(x, y);
                OfficeColor maskPixel = mask.GetPixel(x, y);
                double maskAlpha = maskPixel.A / 255D;
                double coverage = GetMaskFactor(maskPixel, softMask.Mode) + ((1D - maskAlpha) * backdrop);
                result.SetPixel(x, y, OfficeColor.FromRgba(sourcePixel.R, sourcePixel.G, sourcePixel.B, (byte)System.Math.Round(sourcePixel.A * coverage)));
            }
        }
        return result;
    }

    private static double GetMaskFactor(OfficeColor color, OfficeSoftMaskMode mode) {
        double alpha = color.A / 255D;
        if (mode == OfficeSoftMaskMode.Alpha) return alpha;
        return alpha * (((0.3D * color.R) + (0.59D * color.G) + (0.11D * color.B)) / 255D);
    }
}
