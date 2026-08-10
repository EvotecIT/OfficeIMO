using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal readonly struct PdfPagePatternSelection {
    internal PdfPagePatternSelection(
        string name,
        OfficeColor? tint,
        PdfPageColorSpace? baseColorSpace,
        PdfPageTilingPatternResource? tilingPattern) {
        Name = name;
        Tint = tint;
        BaseColorSpace = baseColorSpace;
        TilingPattern = tilingPattern;
    }

    internal string Name { get; }
    internal OfficeColor? Tint { get; }
    internal PdfPageColorSpace? BaseColorSpace { get; }
    internal PdfPageTilingPatternResource? TilingPattern { get; }
}

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
        Matrix2D paintTransform,
        OfficeColor fillColor,
        PdfPageColorSpace fillColorSpace,
        PdfPagePatternSelection? fillPattern,
        PdfPageColorSpace? fillPatternBaseColorSpace,
        double? fillOpacity,
        OfficeColor strokeColor,
        PdfPageColorSpace strokeColorSpace,
        PdfPagePatternSelection? strokePattern,
        PdfPageColorSpace? strokePatternBaseColorSpace,
        double? strokeOpacity,
        double strokeWidth,
        OfficeStrokeDashStyle? strokeDashStyle,
        OfficeStrokeLineCap? strokeLineCap,
        OfficeStrokeLineJoin? strokeLineJoin) {
        Font = font;
        CharacterCode = characterCode;
        Transform = transform;
        ClipPath = clipPath;
        PaintTransform = paintTransform;
        FillColor = fillColor;
        FillColorSpace = fillColorSpace;
        FillPattern = fillPattern;
        FillPatternBaseColorSpace = fillPatternBaseColorSpace;
        FillOpacity = fillOpacity;
        StrokeColor = strokeColor;
        StrokeColorSpace = strokeColorSpace;
        StrokePattern = strokePattern;
        StrokePatternBaseColorSpace = strokePatternBaseColorSpace;
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
    internal Matrix2D PaintTransform { get; }
    internal OfficeColor FillColor { get; }
    internal PdfPageColorSpace FillColorSpace { get; }
    internal PdfPagePatternSelection? FillPattern { get; }
    internal PdfPageColorSpace? FillPatternBaseColorSpace { get; }
    internal double? FillOpacity { get; }
    internal OfficeColor StrokeColor { get; }
    internal PdfPageColorSpace StrokeColorSpace { get; }
    internal PdfPagePatternSelection? StrokePattern { get; }
    internal PdfPageColorSpace? StrokePatternBaseColorSpace { get; }
    internal double? StrokeOpacity { get; }
    internal double StrokeWidth { get; }
    internal OfficeStrokeDashStyle? StrokeDashStyle { get; }
    internal OfficeStrokeLineCap? StrokeLineCap { get; }
    internal OfficeStrokeLineJoin? StrokeLineJoin { get; }
}
