using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Resolves theme-backed table border colors on a rendering copy, preserving source declarations.</summary>
    private static T? ResolveNativeTableBorderColors<T>(T? borders, WordDocument document) where T : OpenXmlCompositeElement {
        if (borders == null || !borders.Elements<W.BorderType>().Any(border => border.ThemeColor != null)) return borders;
        T resolved = (T)borders.CloneNode(true);
        foreach (W.BorderType border in resolved.Elements<W.BorderType>()) {
            string? color = WordRunFormatting.ResolveBorderColorHex(border, document);
            if (color != null) border.Color = color;
        }
        return resolved;
    }
}
