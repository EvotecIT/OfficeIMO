using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal readonly struct PdfPageDrawingEffect {
    public PdfPageDrawingEffect(OfficeBlendMode blendMode, PdfPageSoftMaskResource? softMask) {
        BlendMode = blendMode;
        SoftMask = softMask;
    }

    public static PdfPageDrawingEffect Default => new PdfPageDrawingEffect(OfficeBlendMode.Normal, null);

    public OfficeBlendMode BlendMode { get; }

    public PdfPageSoftMaskResource? SoftMask { get; }

    public bool IsDefault => BlendMode == OfficeBlendMode.Normal && SoftMask == null;

    public PdfPageDrawingEffect Apply(PdfPageGraphicsStateResource resource) => new PdfPageDrawingEffect(
        resource.BlendMode ?? BlendMode,
        resource.HasSoftMask ? resource.SoftMask : SoftMask);
}

internal readonly struct PdfPageDrawingEffectTransition {
    public PdfPageDrawingEffectTransition(double paintOrder, PdfPageDrawingEffect effect) {
        PaintOrder = paintOrder;
        Effect = effect;
    }

    public double PaintOrder { get; }

    public PdfPageDrawingEffect Effect { get; }
}
