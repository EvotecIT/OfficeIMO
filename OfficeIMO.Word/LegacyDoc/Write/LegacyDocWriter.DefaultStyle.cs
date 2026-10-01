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
            var scheme = mainPart.ThemePart?.Theme?.ThemeElements?.FontScheme;
            RunFonts effectiveFonts = fonts == null ? new RunFonts() : (RunFonts)fonts.CloneNode(true);
            effectiveFonts.Ascii = ResolveDefaultThemeFont(scheme, fonts?.AsciiTheme?.Value) ?? fonts?.Ascii?.Value;
            effectiveFonts.HighAnsi = ResolveDefaultThemeFont(scheme, fonts?.HighAnsiTheme?.Value) ?? fonts?.HighAnsi?.Value;
            effectiveFonts.EastAsia = ResolveDefaultThemeFont(scheme, fonts?.EastAsiaTheme?.Value) ?? fonts?.EastAsia?.Value;
            effectiveFonts.ComplexScript = ResolveDefaultThemeFont(scheme, fonts?.ComplexScriptTheme?.Value) ?? fonts?.ComplexScript?.Value;
            string? family = ReadSupportedRunFontFamily(effectiveFonts);
            // A native DOC has no theme-based default font. Materialize the resolved default.
            family = string.IsNullOrWhiteSpace(family) ? "Calibri" : family;
            if (fonts != null) normal.StyleRunProperties.RemoveChild(fonts);
            normal.StyleRunProperties.PrependChild(new RunFonts { Ascii = family, HighAnsi = family });
            return normal;
        }

        private static string? ResolveDefaultThemeFont(A.FontScheme? scheme, ThemeFontValues? selector) {
            if (scheme == null || selector == null) return null;
            string? family = selector.Value switch {
                var value when value == ThemeFontValues.MajorAscii || value == ThemeFontValues.MajorHighAnsi => scheme.MajorFont?.LatinFont?.Typeface?.Value,
                var value when value == ThemeFontValues.MinorAscii || value == ThemeFontValues.MinorHighAnsi => scheme.MinorFont?.LatinFont?.Typeface?.Value,
                var value when value == ThemeFontValues.MajorEastAsia => scheme.MajorFont?.EastAsianFont?.Typeface?.Value,
                var value when value == ThemeFontValues.MinorEastAsia => scheme.MinorFont?.EastAsianFont?.Typeface?.Value,
                var value when value == ThemeFontValues.MajorBidi => scheme.MajorFont?.ComplexScriptFont?.Typeface?.Value,
                var value when value == ThemeFontValues.MinorBidi => scheme.MinorFont?.ComplexScriptFont?.Typeface?.Value,
                _ => throw new NotSupportedException("Native DOC saving cannot resolve the default theme font selector.")
            };
            return string.IsNullOrWhiteSpace(family) ? null : family;
        }
    }
}
