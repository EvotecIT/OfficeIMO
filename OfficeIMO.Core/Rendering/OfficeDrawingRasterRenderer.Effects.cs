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
            TextShapingProvider = canvas.TextShapingProvider,
            TextShapingLanguage = canvas.TextShapingLanguage,
            DiagnosticSink = canvas.DiagnosticSink,
            DiagnosticSource = canvas.DiagnosticSource,
            CancellationToken = cancellationToken
        });
        if (effectGroup.SoftMask != null) {
            layer = ApplySoftMask(
                layer,
                effectGroup.SoftMask,
                scale,
                imageCodec,
                canvas.TextShapingProvider,
                canvas.TextShapingLanguage,
                canvas.DiagnosticSink,
                canvas.DiagnosticSource,
                cancellationToken);
        }
        OfficeTransform transform = effectGroup.Transform;
        var pixelTransform = new OfficeTransform(transform.M11, transform.M12, transform.M21, transform.M22, transform.OffsetX * scale, transform.OffsetY * scale);
        bool interpolate = !ContainsNonInterpolatedImage(effectGroup.InnerDrawing);
        canvas.DrawAffineImage(layer, pixelTransform, effectGroup.Opacity, effectGroup.BlendMode, interpolate);
    }

    private static bool ContainsNonInterpolatedImage(OfficeDrawing drawing) {
        for (int index = 0; index < drawing.Elements.Count; index++) {
            OfficeDrawingElement element = drawing.Elements[index];
            if (element is OfficeDrawingImage { Interpolate: false }) return true;
            if (element is OfficeDrawingGroup group && ContainsNonInterpolatedImage(group.InnerDrawing)) return true;
            if (element is OfficeDrawingEffectGroup effectGroup && ContainsNonInterpolatedImage(effectGroup.InnerDrawing)) return true;
            if (element is OfficeDrawingTilingPattern pattern && ContainsNonInterpolatedImage(pattern.InnerTile)) return true;
        }

        return false;
    }

    private static OfficeRasterImage ApplySoftMask(
        OfficeRasterImage source,
        OfficeDrawingSoftMask softMask,
        double scale,
        IOfficeRasterImageCodec? imageCodec,
        IOfficeTextShapingProvider? textShapingProvider,
        string? textShapingLanguage,
        System.Collections.Generic.ICollection<OfficeImageExportDiagnostic>? diagnosticSink,
        string? diagnosticSource,
        System.Threading.CancellationToken cancellationToken) {
        var maskScene = new OfficeDrawing(source.Width / scale, source.Height / scale);
        maskScene.AddEffectDrawing(softMask.InnerDrawing, softMask.Transform);
        OfficeRasterImage mask = Render(maskScene, new OfficeDrawingRasterRenderOptions {
            Scale = scale,
            ImageCodec = imageCodec,
            TextShapingProvider = textShapingProvider,
            TextShapingLanguage = textShapingLanguage,
            DiagnosticSink = diagnosticSink,
            DiagnosticSource = diagnosticSource,
            CancellationToken = cancellationToken
        });
        var result = new OfficeRasterImage(source.Width, source.Height);
        double backdrop = GetMaskFactor(softMask.BackdropColor, softMask.Mode, softMask.LuminosityStandard);
        for (int y = 0; y < source.Height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int x = 0; x < source.Width; x++) {
                OfficeColor sourcePixel = source.GetPixel(x, y);
                OfficeColor maskPixel = mask.GetPixel(x, y);
                double maskAlpha = maskPixel.A / 255D;
                double coverage = GetMaskFactor(maskPixel, softMask.Mode, softMask.LuminosityStandard) + ((1D - maskAlpha) * backdrop);
                result.SetPixel(x, y, OfficeColor.FromRgba(sourcePixel.R, sourcePixel.G, sourcePixel.B, (byte)System.Math.Round(sourcePixel.A * coverage)));
            }
        }
        return result;
    }

    private static double GetMaskFactor(OfficeColor color, OfficeSoftMaskMode mode, OfficeSoftMaskLuminosityStandard luminosityStandard) {
        double alpha = color.A / 255D;
        if (mode == OfficeSoftMaskMode.Alpha) return alpha;
        double redWeight = luminosityStandard == OfficeSoftMaskLuminosityStandard.PdfDeviceRgb ? 0.3D : 0.2126D;
        double greenWeight = luminosityStandard == OfficeSoftMaskLuminosityStandard.PdfDeviceRgb ? 0.59D : 0.7152D;
        double blueWeight = luminosityStandard == OfficeSoftMaskLuminosityStandard.PdfDeviceRgb ? 0.11D : 0.0722D;
        return alpha * (((redWeight * color.R) + (greenWeight * color.G) + (blueWeight * color.B)) / 255D);
    }
}
