namespace OfficeIMO.Word.Pdf;


public static partial class WordPdfConverterExtensions {
    /// <summary>Resolves theme-backed table border colors on a rendering copy, preserving source declarations.</summary>
    private static OfficeXmlNode ResolveNativeTableBorderColors(OfficeXmlNode borders, WordDocument document) {
        if (borders.IsNone || !borders.Elements().Any(border => IsNativeBorderSideElement(border) && border.WAttr("themeColor") != null)) return borders;
        OfficeXmlNode resolved = borders.Clone();
        foreach (OfficeXmlNode border in resolved.Elements()) {
            if (!IsNativeBorderSideElement(border)) continue;
            string? color = WordConverterNodes.ResolveBorderColorHex(border, document);
            if (color != null) border.SetAttribute(WordXml.W, "color", color);
        }
        return resolved;
    }

    /// <summary>Whether the element is one of the <c>w:</c> border sides a table, cell or paragraph border set holds.</summary>
    private static bool IsNativeBorderSideElement(OfficeXmlNode element) =>
        element.NamespaceId == WordXml.W && element.LocalName is "top" or "left" or "start" or "bottom" or "right" or "end"
            or "insideH" or "insideV" or "tl2br" or "tr2bl" or "between" or "bar";





}
