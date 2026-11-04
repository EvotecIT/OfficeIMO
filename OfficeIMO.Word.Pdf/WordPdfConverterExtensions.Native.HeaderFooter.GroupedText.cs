using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wpg = DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>
    /// Grouped header/footer building blocks print white numbers on a filled shape. Only the text
    /// is drawn, so white text would vanish on the page; it takes the shape's fill colour instead.
    /// </summary>
    private static global::OfficeIMO.OfficeColor? ResolveNativeGroupedHeaderFooterTextColor(WordParagraph paragraph, global::OfficeIMO.OfficeColor? color) {
        if (color is not { } resolved) return color;
        var rgb = resolved.ToNormalizedRgb();
        if (rgb.R < 0.9D || rgb.G < 0.9D || rgb.B < 0.9D) {
            return color;
        }

        W.TextBoxContent? content = paragraph._paragraph?.Ancestors<W.TextBoxContent>().FirstOrDefault();
        Wpg.WordprocessingGroup? group = content?.Ancestors<Wpg.WordprocessingGroup>().FirstOrDefault();
        if (group == null || !content!.Ancestors().Any(parent => parent is W.Header or W.Footer)) {
            return color;
        }

        foreach (Wps.WordprocessingShape shape in group.Descendants<Wps.WordprocessingShape>()) {
            string? fill = shape.GetFirstChild<Wps.ShapeProperties>()?.GetFirstChild<A.SolidFill>()?
                .GetFirstChild<A.RgbColorModelHex>()?.Val?.Value;
            if (ParseNativeColor(fill) is { } fillColor) {
                var fillRgb = fillColor.ToNormalizedRgb();
                if (fillRgb.R < 0.9D || fillRgb.G < 0.9D || fillRgb.B < 0.9D) return fillColor;
            }
        }

        return color;
    }
}
