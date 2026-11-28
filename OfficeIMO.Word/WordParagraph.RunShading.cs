using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>Native run shading fill as an uppercase six-digit hexadecimal string; empty removes it.</summary>
        internal string RunShadingFillValue {
            get {
                // A hyperlink's run, or the paragraph's own run; a content control's run doesn't count here.
                OfficeXmlNode shading = RunPropertiesIgnoringContentControl.Child(W, "shd");
                if (shading.IsNone || shading.Attribute(W, "val") == "nil") return string.Empty;
                string? fill = shading.Attribute(W, "fill");
                return fill != null ? fill.ToUpperInvariant() : string.Empty;
            }
            set {
                OfficeXmlNode properties = EditRunProperties(hyperlinkInContentControl: true);
                if (string.IsNullOrWhiteSpace(value)) {
                    WordXmlEdits.RemoveChild(properties, "shd");
                    return;
                }
                // An existing shading becomes a clear fill of this colour, without theme or pattern colours.
                OfficeXmlNode shading = properties.GetOrAddChild(W, "shd", WRPrView.Order);
                shading.SetAttribute(W, "val", "clear");
                shading.SetAttribute(W, "fill", value.Replace("#", string.Empty).ToUpperInvariant());
                foreach (string attribute in ShadingThemeAttributes) shading.RemoveAttribute(W, attribute);
            }
        }

        private static readonly string[] ShadingThemeAttributes = {
            "themeFill", "themeFillTint", "themeFillShade", "color", "themeColor", "themeTint", "themeShade"
        };

        /// <summary>
        /// Gets or sets the background (run shading fill) color applied to the current text run.
        /// </summary>
        /// <remarks>
        /// Run shading supports arbitrary RGB colors, unlike <see cref="Highlight"/> which is limited
        /// to Word's fixed highlight palette. Setting <see langword="null"/> removes the run shading.
        /// </remarks>
        public Color? BackgroundColor {
            get {
                string fill = RunShadingFillValue;
                return fill.Length == 6 &&
                       OfficeIMO.OfficeColor.TryParseHex(fill, out OfficeIMO.OfficeColor color)
                    ? color
                    : null;
            }
            set => RunShadingFillValue = value?.ToRgbHex() ?? string.Empty;
        }

        /// <summary>
        /// Applies an exact background (run shading fill) color to the current text run.
        /// </summary>
        /// <param name="color">The background color.</param>
        /// <returns>The current paragraph or run wrapper.</returns>
        public WordParagraph SetBackgroundColor(Color color) {
            BackgroundColor = color;
            return this;
        }
    }
}
