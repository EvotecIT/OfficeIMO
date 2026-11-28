using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        // The document being written. Style definitions the writer materializes are detached copies, so fonts inside
        // them find the theme through it.
        [ThreadStatic] private static LegacyDocSourcePackage? t_writingSource;

        /// <summary>The <c>a:fontScheme</c> of the package's theme, or none.</summary>
        private static LegacyDocElement GetThemeFontScheme(LegacyDocSourcePackage? source) {
            if (source == null) return LegacyDocElement.None;
            return source.Theme
                .Child(OfficeXmlNamespaces.DrawingML, "themeElements")
                .Child(OfficeXmlNamespaces.DrawingML, "fontScheme");
        }

        /// <summary>
        /// The concrete font names of a <c>w:rFonts</c> element. Theme selectors resolve through <paramref name="scheme"/>
        /// (the writing package's theme when none), keeping literal names for empty theme slots. Without theme selectors
        /// the element's own names come back unresolved.
        /// </summary>
        private static LegacyDocRunFonts ResolveNativeThemeRunFonts(LegacyDocElement fonts, LegacyDocElement scheme = default) {
            string? asciiTheme = fonts.Attr("asciiTheme");
            string? highAnsiTheme = fonts.Attr("hAnsiTheme");
            string? eastAsiaTheme = fonts.Attr("eastAsiaTheme");
            string? complexScriptTheme = fonts.Attr("cstheme");
            if (asciiTheme == null && highAnsiTheme == null &&
                eastAsiaTheme == null && complexScriptTheme == null) return LegacyDocRunFonts.FromElement(fonts);

            if (scheme.IsNone) {
                scheme = GetThemeFontScheme(t_writingSource);
            }

            // DOC stores concrete font names. Resolve before detaching a style
            // from its package, retaining literal fallbacks for empty theme slots.
            return new LegacyDocRunFonts(
                ResolveDefaultFontSlot(scheme, fonts.Attr("ascii"), asciiTheme, null, null),
                ResolveDefaultFontSlot(scheme, fonts.Attr("hAnsi"), highAnsiTheme, null, null),
                ResolveDefaultFontSlot(scheme, fonts.Attr("eastAsia"), eastAsiaTheme, null, null),
                ResolveDefaultFontSlot(scheme, fonts.Attr("cs"), complexScriptTheme, null, null),
                resolvedTheme: true);
        }

        private static void MaterializeParagraphStyleThemeFonts(Dictionary<string, LegacyDocElement> paragraphStyles, LegacyDocSourcePackage source) {
            LegacyDocElement scheme = GetThemeFontScheme(source);
            foreach (string styleId in paragraphStyles.Keys.ToArray()) {
                LegacyDocElement original = paragraphStyles[styleId];
                LegacyDocElement fonts = original.Child("rPr").Child("rFonts");
                if (fonts.IsNone) continue;
                LegacyDocRunFonts resolved = ResolveNativeThemeRunFonts(fonts, scheme);
                if (!resolved.ResolvedTheme) continue;
                OfficeXmlTree copy = CopyStyleDefinition(original);
                int runFonts = copy.FindChild(copy.FindChild(copy.Root, LegacyDocElement.W, "rPr"), LegacyDocElement.W, "rFonts");
                resolved.ApplyTo(copy, runFonts);
                paragraphStyles[styleId] = new LegacyDocElement(copy, copy.Root);
            }
        }
    }
}
