using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    /// <summary>Writes a literal fallback and retains native theme-slot identity for semantic colors.</summary>
    private void SetRunColor(OfficeColor color) {
        // Resolving a theme colour can read the theme through the SDK, so it runs before the run properties are taken.
        string value = WordRunColor.Resolve(color, _document).ToRgbHex();
        string? themeColor = WordRunColor.ThemeColorToken(color);
        OfficeXmlNode properties = EditRunProperties();
        // A fresh element replaces any previous theme transforms.
        OfficeXmlNode native = WordXmlEdits.ReplaceChild(properties, "color", WRPrView.Order, value);
        if (themeColor != null) native.SetAttribute(W, "themeColor", themeColor);
        SyncPageNumberFormatting(properties, "color");
    }
}
