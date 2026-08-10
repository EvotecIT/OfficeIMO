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
        bool? softMaskEnabled = null,
        PdfPageSoftMaskResource? softMask = null) {
        FillOpacity = fillOpacity;
        StrokeOpacity = strokeOpacity;
        StrokeWidth = strokeWidth;
        StrokeDashStyle = strokeDashStyle;
        StrokeLineCap = strokeLineCap;
        StrokeLineJoin = strokeLineJoin;
        BlendMode = blendMode;
        SoftMaskEnabled = softMaskEnabled;
        SoftMask = softMask;
    }

    public double? FillOpacity { get; }

    public double? StrokeOpacity { get; }

    public double? StrokeWidth { get; }

    public OfficeStrokeDashStyle? StrokeDashStyle { get; }

    public OfficeStrokeLineCap? StrokeLineCap { get; }

    public OfficeStrokeLineJoin? StrokeLineJoin { get; }

    public OfficeBlendMode? BlendMode { get; }

    /// <summary>Null inherits the current mask, false clears it, and true activates a mask.</summary>
    public bool? SoftMaskEnabled { get; }

    public PdfPageSoftMaskResource? SoftMask { get; }
}
