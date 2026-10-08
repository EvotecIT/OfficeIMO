using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static RunFonts ResolveNativeThemeRunFonts(RunFonts fonts, A.FontScheme? scheme = null) {
            if (fonts.AsciiTheme == null && fonts.HighAnsiTheme == null &&
                fonts.EastAsiaTheme == null && fonts.ComplexScriptTheme == null) return fonts;

            if (scheme == null) {
                OpenXmlPartRootElement? root = fonts.Ancestors<OpenXmlPartRootElement>().LastOrDefault();
                MainDocumentPart? main = (root?.OpenXmlPart?.OpenXmlPackage as WordprocessingDocument)?.MainDocumentPart;
                scheme = main?.ThemePart?.Theme?.ThemeElements?.FontScheme;
            }

            // DOC stores concrete font names. Resolve before detaching a style
            // from its package, retaining literal fallbacks for empty theme slots.
            RunFonts resolved = (RunFonts)fonts.CloneNode(true);
            resolved.Ascii = ResolveDefaultFontSlot(scheme, fonts.Ascii?.Value, fonts.AsciiTheme?.Value, null, null);
            resolved.HighAnsi = ResolveDefaultFontSlot(scheme, fonts.HighAnsi?.Value, fonts.HighAnsiTheme?.Value, null, null);
            resolved.EastAsia = ResolveDefaultFontSlot(scheme, fonts.EastAsia?.Value, fonts.EastAsiaTheme?.Value, null, null);
            resolved.ComplexScript = ResolveDefaultFontSlot(scheme, fonts.ComplexScript?.Value, fonts.ComplexScriptTheme?.Value, null, null);
            resolved.AsciiTheme = null;
            resolved.HighAnsiTheme = null;
            resolved.EastAsiaTheme = null;
            resolved.ComplexScriptTheme = null;
            return resolved;
        }

        private static void MaterializeParagraphStyleThemeFonts(Dictionary<string, Style> paragraphStyles, MainDocumentPart main) {
            A.FontScheme? scheme = main.ThemePart?.Theme?.ThemeElements?.FontScheme;
            foreach (string styleId in paragraphStyles.Keys.ToArray()) {
                Style original = paragraphStyles[styleId];
                RunFonts? fonts = original.StyleRunProperties?.RunFonts;
                if (fonts == null) continue;
                RunFonts resolved = ResolveNativeThemeRunFonts(fonts, scheme);
                if (ReferenceEquals(fonts, resolved)) continue;
                Style copy = (Style)original.CloneNode(true);
                copy.StyleRunProperties!.RunFonts = resolved;
                paragraphStyles[styleId] = copy;
            }
        }
    }
}
