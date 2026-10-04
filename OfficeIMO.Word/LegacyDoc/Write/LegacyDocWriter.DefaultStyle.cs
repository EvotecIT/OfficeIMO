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
                    // Font selectors are inherited individually below.
                    if (property is not FontSize && property is not FontSizeComplexScript) continue;
                    if (!normal.StyleRunProperties.ChildElements.Any(child => child.LocalName == property.LocalName))
                        normal.StyleRunProperties.AppendChild(property.CloneNode(true));
                }
            }
            RunFonts? fonts = normal.StyleRunProperties.GetFirstChild<RunFonts>();
            RunFonts? inherited = defaults?.GetFirstChild<RunFonts>();
            var scheme = mainPart.ThemePart?.Theme?.ThemeElements?.FontScheme;
            RunFonts effectiveFonts = new RunFonts {
                Ascii = ResolveDefaultFontSlot(scheme, fonts?.Ascii?.Value, fonts?.AsciiTheme?.Value, inherited?.Ascii?.Value, inherited?.AsciiTheme?.Value),
                HighAnsi = ResolveDefaultFontSlot(scheme, fonts?.HighAnsi?.Value, fonts?.HighAnsiTheme?.Value, inherited?.HighAnsi?.Value, inherited?.HighAnsiTheme?.Value),
                EastAsia = ResolveDefaultFontSlot(scheme, fonts?.EastAsia?.Value, fonts?.EastAsiaTheme?.Value, inherited?.EastAsia?.Value, inherited?.EastAsiaTheme?.Value),
                ComplexScript = ResolveDefaultFontSlot(scheme, fonts?.ComplexScript?.Value, fonts?.ComplexScriptTheme?.Value, inherited?.ComplexScript?.Value, inherited?.ComplexScriptTheme?.Value)
            };
            string? family = ReadSupportedRunFontFamily(effectiveFonts);
            // A native DOC has no theme-based default font. Materialize the resolved default.
            family = string.IsNullOrWhiteSpace(family) ? "Calibri" : family;
            if (fonts != null) normal.StyleRunProperties.RemoveChild(fonts);
            normal.StyleRunProperties.PrependChild(new RunFonts { Ascii = family, HighAnsi = family });
            return normal;
        }

        private static string? ResolveDefaultFontSlot(A.FontScheme? scheme, string? name, ThemeFontValues? selector,
            string? inheritedName, ThemeFontValues? inheritedSelector) =>
            selector != null || !string.IsNullOrWhiteSpace(name)
                ? ResolveDefaultThemeFont(scheme, selector) ?? name
                : ResolveDefaultThemeFont(scheme, inheritedSelector) ?? inheritedName;

        private static void MaterializeDocumentDefaultSpacing(Dictionary<string, Style> paragraphStyles, Styles? styles) {
            SpacingBetweenLines? defaults = styles?.DocDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle?.SpacingBetweenLines;
            if (defaults == null) return;
            // DOC has no docDefaults record. Put the defaults on roots of the style
            // hierarchy; derived styles continue to inherit their base's overrides.
            foreach (string styleId in paragraphStyles.Keys.ToArray()) {
                Style original = paragraphStyles[styleId];
                string? baseId = original.BasedOn?.Val?.Value;
                if (!string.IsNullOrWhiteSpace(baseId) && paragraphStyles.ContainsKey(baseId!)) continue;
                Style style = (Style)original.CloneNode(true);
                StyleParagraphProperties properties = style.StyleParagraphProperties ??= new StyleParagraphProperties();
                SpacingBetweenLines spacing = properties.SpacingBetweenLines ??= new SpacingBetweenLines();
                bool authoredLine = !string.IsNullOrWhiteSpace(spacing.Line?.Value);
                if (!authoredLine && !string.IsNullOrWhiteSpace(defaults.Line?.Value)) {
                    // Word inherits the numeric value and its interpretation
                    // together. A rule without a local value is not an override.
                    spacing.Line = defaults.Line!.Value;
                    spacing.LineRule = defaults.LineRule?.Value ?? LineSpacingRuleValues.Auto;
                }
                foreach (var attribute in defaults.GetAttributes()) {
                    // The line/rule pair is handled above. An authored line
                    // without a rule means automatic spacing.
                    if (attribute.LocalName is "line" or "lineRule") continue;
                    if (!spacing.GetAttributes().Any(existing => existing.LocalName == attribute.LocalName && existing.NamespaceUri == attribute.NamespaceUri)) {
                        spacing.SetAttribute(attribute);
                    }
                }
                paragraphStyles[styleId] = style;
            }
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
