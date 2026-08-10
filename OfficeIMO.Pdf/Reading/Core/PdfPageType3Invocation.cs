using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

[Flags]
internal enum PdfType3PaintChannels {
    None = 0,
    Fill = 1,
    Stroke = 2,
    Both = Fill | Stroke
}

internal readonly struct PdfPagePatternSelection {
    internal PdfPagePatternSelection(
        string name,
        OfficeColor? tint,
        PdfPageColorSpace? baseColorSpace,
        PdfPageTilingPatternResource? tilingPattern,
        PdfPageShadingPatternResource? shadingPattern,
        Matrix2D paintTransform) {
        Name = name;
        Tint = tint;
        BaseColorSpace = baseColorSpace;
        TilingPattern = tilingPattern;
        ShadingPattern = shadingPattern;
        PaintTransform = paintTransform;
    }

    internal string Name { get; }
    internal OfficeColor? Tint { get; }
    internal PdfPageColorSpace? BaseColorSpace { get; }
    internal PdfPageTilingPatternResource? TilingPattern { get; }
    internal PdfPageShadingPatternResource? ShadingPattern { get; }
    internal Matrix2D PaintTransform { get; }

    internal PdfPagePatternSelection Translate(double offsetX, double offsetY, double sourceHeight, double targetHeight) {
        var sourceFlip = new Matrix2D(1D, 0D, 0D, -1D, 0D, sourceHeight);
        var targetFlip = new Matrix2D(1D, 0D, 0D, -1D, 0D, targetHeight);
        Matrix2D translatedPaintTransform = Matrix2D.Multiply(
            targetFlip,
            Matrix2D.Multiply(
                Matrix2D.Translation(-offsetX, -offsetY),
                Matrix2D.Multiply(sourceFlip, PaintTransform)));
        return new PdfPagePatternSelection(
            Name,
            Tint,
            BaseColorSpace,
            TilingPattern,
            ShadingPattern,
            translatedPaintTransform);
    }
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
