using DocumentFormat.OpenXml.Wordprocessing;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>
        /// Gets or sets the exact fill color applied to the current text run.
        /// </summary>
        /// <remarks>
        /// Run shading supports arbitrary RGB colors, unlike Word text highlighting which is limited
        /// to a fixed palette. The value is returned as an uppercase six-digit hexadecimal string.
        /// </remarks>
        public string RunShadingFillColorHex {
            get {
                RunProperties? runProperties = IsHyperLink ? Hyperlink?._runProperties : _runProperties;
                Shading? shading = runProperties?.Shading;
                if (shading?.Val?.Value == ShadingPatternValues.Nil) {
                    return string.Empty;
                }

                string? fill = shading?.Fill?.Value;
                return fill != null ? fill.ToUpperInvariant() : string.Empty;
            }
            set {
                RunProperties runProperties;
                if (IsHyperLink) {
                    var hyperlink = Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }

                if (!string.IsNullOrWhiteSpace(value)) {
                    runProperties.Shading ??= new Shading();
                    runProperties.Shading.Fill = value.Replace("#", string.Empty).ToUpperInvariant();
                    runProperties.Shading.ThemeFill = null;
                    runProperties.Shading.ThemeFillTint = null;
                    runProperties.Shading.ThemeFillShade = null;
                    runProperties.Shading.Color = null;
                    runProperties.Shading.ThemeColor = null;
                    runProperties.Shading.ThemeTint = null;
                    runProperties.Shading.ThemeShade = null;
                    runProperties.Shading.Val = ShadingPatternValues.Clear;
                } else {
                    runProperties.Shading?.Remove();
                }
            }
        }

        /// <summary>
        /// Gets or sets the exact fill color applied to the current text run.
        /// </summary>
        public Color? RunShadingFillColor {
            get {
                string fill = RunShadingFillColorHex;
                return fill.Length == 6 &&
                       OfficeIMO.Drawing.OfficeColor.TryParseHex(fill, out OfficeIMO.Drawing.OfficeColor color)
                    ? color
                    : null;
            }
            set => RunShadingFillColorHex = value?.ToRgbHex() ?? string.Empty;
        }

        /// <summary>
        /// Applies an exact hexadecimal fill color to the current text run.
        /// </summary>
        /// <param name="color">Color in hexadecimal format.</param>
        /// <returns>The current paragraph or run wrapper.</returns>
        public WordParagraph SetRunShadingFillColorHex(string color) {
            RunShadingFillColorHex = color;
            return this;
        }
    }
}
