using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal readonly struct PdfPageGraphicsStateResource {
    public PdfPageGraphicsStateResource(
        double? fillOpacity,
        double? strokeOpacity,
        double? strokeWidth,
        OfficeStrokeDashStyle? strokeDashStyle,
        OfficeStrokeLineCap? strokeLineCap,
        OfficeStrokeLineJoin? strokeLineJoin,
        OfficeBlendMode? blendMode = null,
        bool hasSoftMask = false,
        PdfPageSoftMaskResource? softMask = null,
        bool hasUnsupportedBlendMode = false,
        bool hasUnsupportedSoftMask = false) {
        FillOpacity = fillOpacity;
        StrokeOpacity = strokeOpacity;
        StrokeWidth = strokeWidth;
        StrokeDashStyle = strokeDashStyle;
        StrokeLineCap = strokeLineCap;
        StrokeLineJoin = strokeLineJoin;
        BlendMode = blendMode;
        HasSoftMask = hasSoftMask;
        SoftMask = softMask;
        HasUnsupportedBlendMode = hasUnsupportedBlendMode;
        HasUnsupportedSoftMask = hasUnsupportedSoftMask;
    }

    public double? FillOpacity { get; }

    public double? StrokeOpacity { get; }

    public double? StrokeWidth { get; }

    public OfficeStrokeDashStyle? StrokeDashStyle { get; }

    public OfficeStrokeLineCap? StrokeLineCap { get; }

    public OfficeStrokeLineJoin? StrokeLineJoin { get; }

    public OfficeBlendMode? BlendMode { get; }

    public bool HasSoftMask { get; }

    public PdfPageSoftMaskResource? SoftMask { get; }

    public bool HasUnsupportedBlendMode { get; }

    public bool HasUnsupportedSoftMask { get; }
}
