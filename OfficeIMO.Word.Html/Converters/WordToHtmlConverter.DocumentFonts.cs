using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        /// <summary>
        /// Body and heading fonts come from the document's own styles; the theme's stack in the
        /// default stylesheet is only the fallback when the document names no font.
        /// </summary>
        private static string GetDocumentFontCss(WordDocument document, WordHtmlExportOptions options) {
            if (!string.IsNullOrWhiteSpace(options.FontFamily)) {
                return string.Empty;
            }

            string? bodyFont = ResolveStyleFont(document, "Normal");
            string? headingFont = ResolveStyleFont(document, "Heading1") ?? bodyFont;
            var css = new StringBuilder();
            // The document shells style text through these custom properties with more specific
            // selectors, so override the properties as well as the plain element rules.
            if (!string.IsNullOrWhiteSpace(bodyFont) || !string.IsNullOrWhiteSpace(headingFont)) {
                css.Append(":root{");
                if (!string.IsNullOrWhiteSpace(bodyFont)) {
                    string body = FontFamilyCss(bodyFont!);
                    css.Append("--officeimo-font-body:").Append(body).Append(";--office-font-body:").Append(body).Append(';');
                }
                if (!string.IsNullOrWhiteSpace(headingFont)) {
                    string heading = FontFamilyCss(headingFont!);
                    css.Append("--officeimo-font-heading:").Append(heading).Append(";--office-font-heading:").Append(heading).Append(';');
                }
                css.Append('}');
            }
            if (!string.IsNullOrWhiteSpace(bodyFont)) {
                css.Append("body{font-family:").Append(FontFamilyCss(bodyFont!)).Append('}');
            }
            if (!string.IsNullOrWhiteSpace(headingFont)) {
                css.Append("h1,h2,h3,h4,h5,h6,nav.toc .toc-title{font-family:").Append(FontFamilyCss(headingFont!)).Append('}');
            }
            return css.ToString();
        }

        private static string? ResolveStyleFont(WordDocument document, string styleId) {
            var mainPart = document._wordprocessingDocument?.MainDocumentPart;
            Styles? styles = mainPart?.StyleDefinitionsPart?.Styles;
            if (styles == null) {
                return null;
            }

            foreach (Style style in WordRunFormatting.EnumerateStyleChain(styles, styleId, StyleValues.Paragraph)) {
                if (!string.Equals(style.StyleId?.Value, styleId, StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(styleId, "Normal", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(style.StyleId?.Value, "Normal", StringComparison.OrdinalIgnoreCase)) {
                    // A heading that names no font of its own uses the body font.
                    return null;
                }
                string? font = ResolveRunFonts(style.StyleRunProperties?.RunFonts, mainPart);
                if (font != null) {
                    return font;
                }
            }

            if (!string.Equals(styleId, "Normal", StringComparison.OrdinalIgnoreCase)) {
                return null;
            }

            RunFonts? defaults = styles.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle?.RunFonts;
            return ResolveRunFonts(defaults, mainPart);
        }

        private static string? ResolveRunFonts(RunFonts? fonts, DocumentFormat.OpenXml.Packaging.MainDocumentPart? mainPart) {
            if (fonts == null) {
                return null;
            }
            if (!string.IsNullOrWhiteSpace(fonts.Ascii?.Value)) {
                return fonts.Ascii!.Value!.Trim();
            }
            ThemeFontValues? theme = fonts.AsciiTheme?.Value ?? fonts.HighAnsiTheme?.Value;
            if (theme == null) {
                return string.IsNullOrWhiteSpace(fonts.HighAnsi?.Value) ? null : fonts.HighAnsi!.Value!.Trim();
            }

            A.FontScheme? scheme = mainPart?.ThemePart?.Theme?.ThemeElements?.FontScheme;
            bool major = theme == ThemeFontValues.MajorAscii || theme == ThemeFontValues.MajorHighAnsi ||
                theme == ThemeFontValues.MajorBidi || theme == ThemeFontValues.MajorEastAsia;
            string? typeface = major
                ? scheme?.MajorFont?.LatinFont?.Typeface?.Value
                : scheme?.MinorFont?.LatinFont?.Typeface?.Value;
            return string.IsNullOrWhiteSpace(typeface) ? null : typeface!.Trim();
        }

        private static string? GetFirstHeadingText(WordDocument document) {
            foreach (WordParagraph paragraph in document.Paragraphs.Take(200)) {
                if (string.Equals(paragraph.StyleId, "Title", StringComparison.OrdinalIgnoreCase) ||
                    (paragraph.Style.HasValue && WordHeadingStyleMapper.GetLevelForHeadingStyle(paragraph.Style.Value) == 1)) {
                    string text = paragraph.Text.Trim();
                    if (text.Length > 0) {
                        return text.Length > 200 ? text.Substring(0, 200) : text;
                    }
                }
            }
            return null;
        }
    }
}
