using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static Style CreateDefaultParagraphStyle(MainDocumentPart mainPart,
            IReadOnlyDictionary<string, Style> paragraphStyles, Styles? styles) {
            Style normal = paragraphStyles.TryGetValue("Normal", out Style? existing)
                ? (Style)existing.CloneNode(true)
                : new Style { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true };
            normal.StyleRunProperties ??= new StyleRunProperties();
            var defaults = styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
            if (defaults != null) {
                foreach (var property in defaults.ChildElements) {
                    // Only materialize the default font and size here. Other script-specific
                    // defaults retain their existing native-writer support boundary.
                    if (property is not RunFonts && property is not FontSize && property is not FontSizeComplexScript) continue;
                    if (!normal.StyleRunProperties.ChildElements.Any(child => child.LocalName == property.LocalName))
                        normal.StyleRunProperties.AppendChild(property.CloneNode(true));
                }
            }
            RunFonts? fonts = normal.StyleRunProperties.GetFirstChild<RunFonts>();
            string? family = fonts?.Ascii?.Value ?? fonts?.HighAnsi?.Value;
            if (family == null && fonts?.AsciiTheme?.Value is { } theme) {
                var scheme = mainPart.ThemePart?.Theme?.ThemeElements?.FontScheme;
                family = theme == ThemeFontValues.MajorHighAnsi
                    ? scheme?.MajorFont?.GetFirstChild<A.LatinFont>()?.Typeface?.Value
                    : scheme?.MinorFont?.GetFirstChild<A.LatinFont>()?.Typeface?.Value;
            }
            // A native DOC has no theme-based default font. Materialize the resolved default.
            family = string.IsNullOrWhiteSpace(family) ? "Calibri" : family;
            if (fonts != null) normal.StyleRunProperties.RemoveChild(fonts);
            normal.StyleRunProperties.PrependChild(new RunFonts { Ascii = family, HighAnsi = family });
            return normal;
        }
    }
}
