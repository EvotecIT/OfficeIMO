using System.Text;
using OfficeIMO.OpenXml.Internal;

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

            string? bodyFont = ResolveStyleFont(document, document.ResolveBuiltInStyleId("Normal"));
            string? headingFont = ResolveStyleFont(document, document.ResolveBuiltInStyleId("Heading1")) ?? bodyFont;
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
            WordStyleCatalog styles = WordStyleCatalog.For(document);
            if (!document.HasStylesPart()) {
                return null;
            }

            // A localized template stores Normal under a translated identifier, such as Standard.
            string normalStyleId = document.ResolveBuiltInStyleId("Normal");
            bool isNormal = string.Equals(styleId, normalStyleId, StringComparison.Ordinal);
            foreach (OfficeXmlNode style in styles.StyleChainNodes(styleId, "paragraph")) {
                string? chainStyleId = style.WAttr("styleId");
                if (!string.Equals(chainStyleId, styleId, StringComparison.Ordinal) &&
                    !isNormal &&
                    string.Equals(chainStyleId, normalStyleId, StringComparison.Ordinal)) {
                    // A heading that names no font of its own uses the body font.
                    return null;
                }
                string? font = ResolveRunFonts(style.WChild("rPr").WChild("rFonts"), document);
                if (font != null) {
                    return font;
                }
            }

            if (!isNormal) {
                return null;
            }

            OfficeXmlNode defaults = styles.DocumentDefaultsNode().WChild("rPrDefault").WChild("rPr").WChild("rFonts");
            return ResolveRunFonts(defaults, document);
        }

        private static string? ResolveRunFonts(OfficeXmlNode fonts, WordDocument document) {
            if (fonts.IsNone) {
                return null;
            }
            string? ascii = fonts.WAttr("ascii");
            string? highAnsi = fonts.WAttr("hAnsi");
            // A theme font selection takes precedence over the family name cached beside it, as in Word.
            string? themeFont = fonts.WAttr("asciiTheme") ?? fonts.WAttr("hAnsiTheme");
            if (themeFont == null) {
                if (!string.IsNullOrWhiteSpace(ascii)) return ascii!.Trim();
                return string.IsNullOrWhiteSpace(highAnsi) ? null : highAnsi!.Trim();
            }

            OfficeXmlNode scheme = document.ThemeNode().Child(WordXml.A, "themeElements").Child(WordXml.A, "fontScheme");
            bool major = themeFont == "majorAscii" || themeFont == "majorHAnsi" ||
                themeFont == "majorBidi" || themeFont == "majorEastAsia";
            string? typeface = scheme.Child(WordXml.A, major ? "majorFont" : "minorFont").Child(WordXml.A, "latin").Attr("typeface");
            if (string.IsNullOrWhiteSpace(typeface)) typeface = ascii ?? highAnsi;
            return string.IsNullOrWhiteSpace(typeface) ? null : typeface!.Trim();
        }

        private static string? GetFirstHeadingText(WordDocument document) {
            foreach (WordParagraph paragraph in document.Paragraphs.Take(200)) {
                if (string.Equals(paragraph.BuiltInStyleId, "Title", StringComparison.Ordinal) ||
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
