using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>Native paragraph shading fill (hexadecimal or <c>auto</c>); empty removes the shading.</summary>
        internal string ShadingFillValue {
            get {
                string? fill = ParagraphPropertiesNode.Child(W, "shd").Attribute(W, "fill");
                return fill != null ? fill.ToUpperInvariant() : string.Empty;
            }
            set {
                OfficeXmlNode properties = EditParagraphProperties();
                if (value == string.Empty) {
                    WordXmlEdits.RemoveChild(properties, "shd");
                    return;
                }
                OfficeXmlNode shading = properties.GetOrAddChild(W, "shd", WPPrView.Order);
                if (!shading.HasAttribute(W, "val")) shading.SetAttribute(W, "val", "clear");
                shading.SetAttribute(W, "fill", Helpers.NormalizeOpenXmlColor(value));
                // An explicit fill replaces any theme fill selector loaded with the paragraph.
                shading.RemoveAttribute(W, "themeFill");
                shading.RemoveAttribute(W, "themeFillTint");
                shading.RemoveAttribute(W, "themeFillShade");
            }
        }

        /// <summary>
        /// Get or set the paragraph shading fill color. Setting <see langword="null"/> removes the shading.
        /// </summary>
        public Color? ShadingFillColor {
            get {
                string fill = ShadingFillValue;
                return fill.Length > 0 && OfficeIMO.OfficeColor.TryParseHex(fill, out OfficeIMO.OfficeColor color) ? color : null;
            }
            set => ShadingFillValue = value?.ToRgbHex() ?? string.Empty;
        }

        /// <summary>
        /// Get or set the paragraph shading pattern.
        /// </summary>
        public WordShadingPattern? ShadingPattern {
            get {
                string? token = ParagraphPropertiesNode.Child(W, "shd").Attribute(W, "val");
                return token == null ? null : WordXmlEdits.Token(WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordShadingPattern pattern), pattern, token);
            }
            set {
                OfficeXmlNode properties = EditParagraphProperties();
                if (value != null) properties.GetOrAddChild(W, "shd", WPPrView.Order).SetAttribute(W, "val", value.Value.ToOpenXmlToken());
                else WordXmlEdits.RemoveChild(properties, "shd");
            }
        }
    }
}
