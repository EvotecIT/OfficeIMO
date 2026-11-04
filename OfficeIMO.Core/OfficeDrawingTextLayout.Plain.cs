using System;

namespace OfficeIMO.Drawing;

internal static partial class OfficeDrawingTextLayout {
    /// <summary>Measures an ordinary Drawing text frame using the same font and fit path as image export.</summary>
    internal static OfficeTextBlockLayout CreateWithRasterMetrics(OfficeDrawingText text, double contentWidth,
        double contentHeight, OfficeRasterCanvas textMetrics) {
        double sourceFontSize = text.Font.Size > 0 ? text.Font.Size : 10D;
        double fontSize = sourceFontSize * text.BaselineScale;
        double lineHeightFactor = ResolveLineHeightFactor(text.LineHeight, fontSize);
        double minimumFontSize = Math.Min(6D, fontSize);
        Func<string?, double, double> measure = (value, size) =>
            textMetrics.MeasureText(value, size, text.Font.FamilyName, text.Font.Style);
        return text.StackedText
            ? OfficeTextLayoutEngine.LayoutStackedTextBlockCore(
                text.RasterText,
                fontSize,
                contentWidth,
                contentHeight,
                lineHeightFactor,
                minimumFontSize,
                measure,
                text.ShrinkToFit,
                (value, size) => textMetrics.MeasureTextPaintBounds(value, size, text.Font.FamilyName, text.Font.Style))
            : text.ShrinkToFit && text.WrapText
            ? OfficeTextLayoutEngine.FitWrappedTextCore(
                text.RasterText,
                fontSize,
                contentWidth,
                contentHeight,
                lineHeightFactor,
                minimumFontSize,
                measure,
                text.ParagraphIndent,
                (value, size) => textMetrics.MeasureTextPaintBounds(value, size, text.Font.FamilyName, text.Font.Style))
            : OfficeTextLayoutEngine.LayoutTextBlock(
                text.RasterText,
                fontSize,
                contentWidth,
                contentHeight,
                lineHeightFactor,
                minimumFontSize,
                measure,
                wrap: text.WrapText,
                shrinkToFit: text.ShrinkToFit,
                paragraphIndent: text.ParagraphIndent);
    }
}
