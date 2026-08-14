using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal readonly struct PdfPageDrawingEffect {
    private PdfPageDrawingEffect(
        OfficeBlendMode blendMode,
        PdfPageSoftMaskResource? softMask,
        bool hasBlendMode,
        bool hasSoftMask,
        Matrix2D? softMaskTransform,
        OfficeIccRenderingIntent renderingIntent,
        bool hasRenderingIntent) {
        BlendMode = blendMode;
        SoftMask = softMask;
        HasBlendMode = hasBlendMode;
        HasSoftMask = hasSoftMask;
        SoftMaskTransform = softMaskTransform;
        RenderingIntent = renderingIntent;
        HasRenderingIntent = hasRenderingIntent;
    }

    public static PdfPageDrawingEffect Default => new PdfPageDrawingEffect(OfficeBlendMode.Normal, null, false, false, null, OfficeIccRenderingIntent.RelativeColorimetric, false);

    public OfficeBlendMode BlendMode { get; }

    public PdfPageSoftMaskResource? SoftMask { get; }

    internal bool HasBlendMode { get; }

    internal bool HasSoftMask { get; }

    internal Matrix2D? SoftMaskTransform { get; }

    internal OfficeIccRenderingIntent RenderingIntent { get; }

    internal bool HasRenderingIntent { get; }

    public bool IsDefault => BlendMode == OfficeBlendMode.Normal && SoftMask == null;

    public PdfPageDrawingEffect Apply(PdfPageGraphicsStateResource resource) => new PdfPageDrawingEffect(
        resource.BlendMode ?? BlendMode,
        resource.SoftMaskEnabled.HasValue
            ? resource.SoftMaskEnabled.Value ? resource.SoftMask : null
            : SoftMask,
        HasBlendMode || resource.BlendMode.HasValue,
        HasSoftMask || resource.SoftMaskEnabled.HasValue,
        resource.SoftMaskEnabled.HasValue ? null : SoftMaskTransform,
        resource.RenderingIntent ?? RenderingIntent,
        HasRenderingIntent || resource.RenderingIntent.HasValue);

    internal PdfPageDrawingEffect OverlayOn(PdfPageDrawingEffect inherited) => new PdfPageDrawingEffect(
        HasBlendMode ? BlendMode : inherited.BlendMode,
        HasSoftMask ? SoftMask : inherited.SoftMask,
        inherited.HasBlendMode || HasBlendMode,
        inherited.HasSoftMask || HasSoftMask,
        HasSoftMask ? SoftMaskTransform : inherited.SoftMaskTransform,
        HasRenderingIntent ? RenderingIntent : inherited.RenderingIntent,
        inherited.HasRenderingIntent || HasRenderingIntent);

    internal PdfPageDrawingEffect WithSoftMaskTransform(Matrix2D transform) => new PdfPageDrawingEffect(
        BlendMode,
        SoftMask,
        HasBlendMode,
        HasSoftMask,
        SoftMask == null ? null : transform,
        RenderingIntent,
        HasRenderingIntent);

    internal PdfPageDrawingEffect WithRenderingIntent(OfficeIccRenderingIntent renderingIntent) => new PdfPageDrawingEffect(
        BlendMode,
        SoftMask,
        HasBlendMode,
        HasSoftMask,
        SoftMaskTransform,
        renderingIntent,
        true);

    internal PdfPageDrawingEffect WithEffectiveRenderingIntent(OfficeIccRenderingIntent renderingIntent) => new PdfPageDrawingEffect(
        BlendMode,
        SoftMask,
        HasBlendMode,
        HasSoftMask,
        SoftMaskTransform,
        renderingIntent,
        HasRenderingIntent);
}

internal readonly struct PdfPageDrawingEffectTransition {
    public PdfPageDrawingEffectTransition(
        double paintOrder,
        PdfPageDrawingEffect effect,
        PdfContentOrderKey? contentOrderKey = null,
        int contentNestingDepth = 0) {
        PaintOrder = paintOrder;
        Effect = effect;
        ContentOrderKey = contentOrderKey;
        ContentNestingDepth = contentNestingDepth;
    }

    public double PaintOrder { get; }

    public PdfPageDrawingEffect Effect { get; }

    internal PdfContentOrderKey? ContentOrderKey { get; }

    internal int ContentNestingDepth { get; }
}
