using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Placement geometry for one image XObject invocation on a PDF page.
/// </summary>
public sealed class PdfImagePlacement {
    internal PdfImagePlacement(
        int pageNumber,
        string resourceName,
        int objectNumber,
        int directStreamIdentity,
        double a,
        double b,
        double c,
        double d,
        double e,
        double f,
        double x,
        double y,
        double width,
        double height,
        PdfPageClipPath? clipPath = null,
        OfficeColor? imageMaskColor = null,
        double? imageOpacity = null,
        PdfStream? inlineImageStream = null,
        PdfDictionary? inlineImageResources = null,
        double paintOrder = 0D,
        OfficeIccRenderingIntent renderingIntent = OfficeIccRenderingIntent.RelativeColorimetric,
        OfficeBlendMode? blendMode = null,
        bool hasUnsupportedBlendMode = false,
        bool hasSoftMask = false,
        bool hasAuthoredRenderingIntent = false,
        PdfContentOrderKey? contentOrderKey = null,
        PdfPagePatternSelection? fillPattern = null,
        PdfDictionary? effectiveResources = null,
        bool requireExactProjection = false,
        bool isHiddenOptionalContent = false) {
        PageNumber = pageNumber;
        ResourceName = resourceName;
        ObjectNumber = objectNumber;
        DirectStreamIdentity = directStreamIdentity;
        A = a;
        B = b;
        C = c;
        D = d;
        E = e;
        F = f;
        X = x;
        Y = y;
        Width = width;
        Height = height;
        ClipPath = clipPath;
        Clip = clipPath.HasValue ? new PdfImageClipInfo(clipPath.Value) : null;
        ImageMaskColor = imageMaskColor ?? OfficeColor.Black;
        ImageOpacity = imageOpacity;
        InlineImageStream = inlineImageStream;
        InlineImageResources = inlineImageResources;
        PaintOrder = paintOrder;
        RenderingIntent = renderingIntent;
        BlendMode = blendMode;
        HasUnsupportedBlendMode = hasUnsupportedBlendMode;
        HasSoftMask = hasSoftMask;
        HasAuthoredRenderingIntent = hasAuthoredRenderingIntent;
        ContentOrderKey = contentOrderKey;
        FillPattern = fillPattern;
        EffectiveResources = effectiveResources;
        RequireExactProjection = requireExactProjection;
        IsHiddenOptionalContent = isHiddenOptionalContent;
    }

    /// <summary>One-based source page number containing the image invocation.</summary>
    public int PageNumber { get; }

    /// <summary>Image resource name from the page or form XObject resource dictionary.</summary>
    public string ResourceName { get; }

    /// <summary>PDF object number for the image stream, or 0 when the image is direct.</summary>
    public int ObjectNumber { get; }

    /// <summary>Runtime identity for a direct image stream, or 0 when the image is indirect.</summary>
    internal int DirectStreamIdentity { get; }

    /// <summary>Current transformation matrix A component at the image invocation.</summary>
    public double A { get; }

    /// <summary>Current transformation matrix B component at the image invocation.</summary>
    public double B { get; }

    /// <summary>Current transformation matrix C component at the image invocation.</summary>
    public double C { get; }

    /// <summary>Current transformation matrix D component at the image invocation.</summary>
    public double D { get; }

    /// <summary>Current transformation matrix E translation component at the image invocation.</summary>
    public double E { get; }

    /// <summary>Current transformation matrix F translation component at the image invocation.</summary>
    public double F { get; }

    /// <summary>Left edge of the placement bounding box in PDF points.</summary>
    public double X { get; }

    /// <summary>Bottom edge of the placement bounding box in PDF points.</summary>
    public double Y { get; }

    /// <summary>Bounding-box width in PDF points.</summary>
    public double Width { get; }

    /// <summary>Bounding-box height in PDF points.</summary>
    public double Height { get; }

    internal PdfPageClipPath? ClipPath { get; }

    /// <summary>Effective clipping path applied to this image placement, when present.</summary>
    public PdfImageClipInfo? Clip { get; }

    /// <summary>Paint color applied to a stencil image mask.</summary>
    public OfficeColor ImageMaskColor { get; }

    /// <summary>Authored or inherited nondefault opacity, or null when the effective opacity is one.</summary>
    public double? ImageOpacity { get; }

    /// <summary>Authored or inherited nondefault opacity, or null when the effective opacity is one.</summary>
    public double? AuthoredOpacity => ImageOpacity;

    /// <summary>Effective image opacity from zero through one.</summary>
    public double Opacity => ImageOpacity ?? 1D;

    internal PdfStream? InlineImageStream { get; }

    internal PdfDictionary? InlineImageResources { get; }

    /// <summary>Stable page-content paint order used to interleave images with other recovered primitives.</summary>
    public double PaintOrder { get; }

    /// <summary>Effective ICC rendering intent for the image placement.</summary>
    public OfficeIccRenderingIntent RenderingIntent { get; }

    /// <summary>Authored or inherited supported blend mode, or null for the normal default.</summary>
    public OfficeBlendMode? BlendMode { get; }

    /// <summary>Authored or inherited supported blend mode, or null for the normal default.</summary>
    public OfficeBlendMode? AuthoredBlendMode => BlendMode;

    /// <summary>Effective supported blend mode.</summary>
    public OfficeBlendMode EffectiveBlendMode => BlendMode ?? OfficeBlendMode.Normal;

    /// <summary>True when the source declared a blend mode that could not be represented.</summary>
    public bool HasUnsupportedBlendMode { get; }

    /// <summary>True when a soft-mask graphics state applies to the placement.</summary>
    public bool HasSoftMask { get; }

    /// <summary>True when the source explicitly selected a rendering intent.</summary>
    public bool HasAuthoredRenderingIntent { get; }

    internal string? SourceDocumentIdentity { get; set; }

    internal PdfContentOrderKey? ContentOrderKey { get; }

    internal PdfPagePatternSelection? FillPattern { get; }

    internal PdfDictionary? EffectiveResources { get; }

    internal bool RequireExactProjection { get; }

    internal bool IsHiddenOptionalContent { get; }

    internal PdfImagePlacement WithPaintOrder(double paintOrder) =>
        Copy(ImageMaskColor, paintOrder);

    internal PdfImagePlacement WithImageMaskColor(OfficeColor imageMaskColor) =>
        Copy(imageMaskColor, PaintOrder);

    internal PdfImagePlacement WithContentOrderKey(PdfContentOrderKey contentOrderKey) =>
        new PdfImagePlacement(
            PageNumber, ResourceName, ObjectNumber, DirectStreamIdentity,
            A, B, C, D, E, F, X, Y, Width, Height, ClipPath,
            ImageMaskColor, ImageOpacity, InlineImageStream, InlineImageResources,
            PaintOrder,
            renderingIntent: RenderingIntent,
            blendMode: BlendMode,
            hasUnsupportedBlendMode: HasUnsupportedBlendMode,
            hasSoftMask: HasSoftMask,
            hasAuthoredRenderingIntent: HasAuthoredRenderingIntent,
            contentOrderKey: contentOrderKey,
            fillPattern: FillPattern,
            effectiveResources: EffectiveResources,
            requireExactProjection: RequireExactProjection,
            isHiddenOptionalContent: IsHiddenOptionalContent) {
            SourceDocumentIdentity = this.SourceDocumentIdentity
        };

    internal PdfImagePlacement WithExactProjection() =>
        new PdfImagePlacement(
            PageNumber, ResourceName, ObjectNumber, DirectStreamIdentity,
            A, B, C, D, E, F, X, Y, Width, Height, ClipPath,
            ImageMaskColor, ImageOpacity, InlineImageStream, InlineImageResources,
            PaintOrder,
            renderingIntent: RenderingIntent,
            blendMode: BlendMode,
            hasUnsupportedBlendMode: HasUnsupportedBlendMode,
            hasSoftMask: HasSoftMask,
            hasAuthoredRenderingIntent: HasAuthoredRenderingIntent,
            contentOrderKey: ContentOrderKey,
            fillPattern: FillPattern,
            effectiveResources: EffectiveResources,
            requireExactProjection: true,
            isHiddenOptionalContent: IsHiddenOptionalContent) {
            SourceDocumentIdentity = this.SourceDocumentIdentity
        };

    internal PdfImagePlacement WithHiddenOptionalContent(bool isHiddenOptionalContent) =>
        new PdfImagePlacement(
            PageNumber, ResourceName, ObjectNumber, DirectStreamIdentity,
            A, B, C, D, E, F, X, Y, Width, Height, ClipPath,
            ImageMaskColor, ImageOpacity, InlineImageStream, InlineImageResources,
            PaintOrder,
            renderingIntent: RenderingIntent,
            blendMode: BlendMode,
            hasUnsupportedBlendMode: HasUnsupportedBlendMode,
            hasSoftMask: HasSoftMask,
            hasAuthoredRenderingIntent: HasAuthoredRenderingIntent,
            contentOrderKey: ContentOrderKey,
            fillPattern: FillPattern,
            effectiveResources: EffectiveResources,
            requireExactProjection: RequireExactProjection,
            isHiddenOptionalContent: isHiddenOptionalContent) {
            SourceDocumentIdentity = this.SourceDocumentIdentity
        };

    private PdfImagePlacement Copy(OfficeColor imageMaskColor, double paintOrder) =>
        new PdfImagePlacement(
            PageNumber,
            ResourceName,
            ObjectNumber,
            DirectStreamIdentity,
            A,
            B,
            C,
            D,
            E,
            F,
            X,
            Y,
            Width,
            Height,
            ClipPath,
            imageMaskColor,
            ImageOpacity,
            InlineImageStream,
            InlineImageResources,
            paintOrder,
            renderingIntent: RenderingIntent,
            blendMode: BlendMode,
            hasUnsupportedBlendMode: HasUnsupportedBlendMode,
            hasSoftMask: HasSoftMask,
            hasAuthoredRenderingIntent: HasAuthoredRenderingIntent,
            contentOrderKey: ContentOrderKey,
            fillPattern: FillPattern,
            effectiveResources: EffectiveResources,
            requireExactProjection: RequireExactProjection,
            isHiddenOptionalContent: IsHiddenOptionalContent) {
            SourceDocumentIdentity = this.SourceDocumentIdentity
        };

    /// <summary>True when the placement matrix is axis-aligned within a small tolerance.</summary>
    public bool IsAxisAligned => Math.Abs(B) <= 0.001D && Math.Abs(C) <= 0.001D;
}
