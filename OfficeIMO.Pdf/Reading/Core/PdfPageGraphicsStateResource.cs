using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal readonly struct PdfPageGraphicsStateResource {
    public PdfPageGraphicsStateResource(
        double? fillOpacity,
        double? strokeOpacity,
        double? strokeWidth,
        OfficeStrokeDashStyle? strokeDashStyle,
        OfficeStrokeLineCap? strokeLineCap,
        OfficeStrokeLineJoin? strokeLineJoin) {
        FillOpacity = fillOpacity;
        StrokeOpacity = strokeOpacity;
        StrokeWidth = strokeWidth;
        StrokeDashStyle = strokeDashStyle;
        StrokeLineCap = strokeLineCap;
        StrokeLineJoin = strokeLineJoin;
    }

    public double? FillOpacity { get; }

    public double? StrokeOpacity { get; }

    public double? StrokeWidth { get; }

    public OfficeStrokeDashStyle? StrokeDashStyle { get; }

    public OfficeStrokeLineCap? StrokeLineCap { get; }

    public OfficeStrokeLineJoin? StrokeLineJoin { get; }
}
