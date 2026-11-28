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

        OfficeXmlNode content = paragraph.Node().Ancestor(WordXml.W, "txbxContent");
        OfficeXmlNode group = content.Ancestors().FirstOrDefault(ancestor => ancestor.Is(WordXml.Wpg, "wgp"));
        if (group.IsNone || !content.Ancestors().Any(parent => parent.IsW("hdr") || parent.IsW("ftr"))) {
            return color;
        }

        foreach (OfficeXmlNode shape in group.DescendantsSkippingExtLst(WordXml.Wps, "wsp")) {
            string? fill = shape.Child(WordXml.Wps, "spPr").Child(WordXml.A, "solidFill")
                .Child(WordXml.A, "srgbClr").Attr("val");
            if (ParseNativeColor(fill) is { } fillColor) {
                var fillRgb = fillColor.ToNormalizedRgb();
                if (fillRgb.R < 0.9D || fillRgb.G < 0.9D || fillRgb.B < 0.9D) return fillColor;
            }
        }

        return color;
    }
}
