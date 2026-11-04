using DocumentFormat.OpenXml.Wordprocessing;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>Native run shading fill as an uppercase six-digit hexadecimal string; empty removes it.</summary>
        internal string RunShadingFillValue {
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
