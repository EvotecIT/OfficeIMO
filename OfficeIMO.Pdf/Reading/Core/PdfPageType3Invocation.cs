using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal readonly struct PdfPageType3TextInvocation {
    internal PdfPageType3TextInvocation(IReadOnlyList<PdfPageType3GlyphInvocation> glyphs, double paintOrder, int sourceOperatorIndex) {
        Glyphs = glyphs;
        PaintOrder = paintOrder;
        SourceOperatorIndex = sourceOperatorIndex;
    }

    internal IReadOnlyList<PdfPageType3GlyphInvocation> Glyphs { get; }

    internal double PaintOrder { get; }

    internal int SourceOperatorIndex { get; }
}

internal readonly struct PdfPageType3GlyphInvocation {
    internal PdfPageType3GlyphInvocation(
        PdfFontResource font,
        byte characterCode,
        Matrix2D transform,
        PdfPageClipPath? clipPath,
        OfficeColor fillColor,
        PdfPageColorSpace fillColorSpace,
        double? fillOpacity,
        OfficeColor strokeColor,
        PdfPageColorSpace strokeColorSpace,
        double? strokeOpacity,
        double strokeWidth,
        OfficeStrokeDashStyle? strokeDashStyle,
        OfficeStrokeLineCap? strokeLineCap,
        OfficeStrokeLineJoin? strokeLineJoin) {
        Font = font;
        CharacterCode = characterCode;
        Transform = transform;
        ClipPath = clipPath;
        FillColor = fillColor;
        FillColorSpace = fillColorSpace;
        FillOpacity = fillOpacity;
        StrokeColor = strokeColor;
        StrokeColorSpace = strokeColorSpace;
        StrokeOpacity = strokeOpacity;
        StrokeWidth = strokeWidth;
        StrokeDashStyle = strokeDashStyle;
        StrokeLineCap = strokeLineCap;
        StrokeLineJoin = strokeLineJoin;
    }

    internal PdfFontResource Font { get; }
    internal byte CharacterCode { get; }
    internal Matrix2D Transform { get; }
    internal PdfPageClipPath? ClipPath { get; }
    internal OfficeColor FillColor { get; }
    internal PdfPageColorSpace FillColorSpace { get; }
    internal double? FillOpacity { get; }
    internal OfficeColor StrokeColor { get; }
    internal PdfPageColorSpace StrokeColorSpace { get; }
    internal double? StrokeOpacity { get; }
    internal double StrokeWidth { get; }
    internal OfficeStrokeDashStyle? StrokeDashStyle { get; }
    internal OfficeStrokeLineCap? StrokeLineCap { get; }
    internal OfficeStrokeLineJoin? StrokeLineJoin { get; }
}
