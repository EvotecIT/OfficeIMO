using DocumentFormat.OpenXml.Wordprocessing;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>
        /// Get or set the paragraph shading fill color using hexadecimal value.
        /// </summary>
        public string ShadingFillColorHex {
            get {
                var fill = _paragraphProperties?.Shading?.Fill?.Value;
                return fill != null ? fill.ToLowerInvariant() : string.Empty;
            }
            set {
                var props = _paragraph.ParagraphProperties ??= new ParagraphProperties();
                if (value != string.Empty) {
                    var color = value.Replace("#", string.Empty).ToLowerInvariant();
                    props.Shading ??= new Shading();
                    props.Shading.Fill = color;
                    props.Shading.Val ??= ShadingPatternValues.Clear;
                } else {
                    props.Shading?.Remove();
                }
            }
        }

        /// <summary>
        /// Get or set the paragraph shading fill color.
        /// </summary>
        public Color? ShadingFillColor {
            get {
                if (ShadingFillColorHex != string.Empty) {
                    return Helpers.ParseColor(ShadingFillColorHex);
                }

                return null;
            }
            set {
                if (value != null) {
                    ShadingFillColorHex = value.Value.ToHexColor();
                }
            }
        }

        /// <summary>
        /// Get or set the paragraph shading pattern.
        /// </summary>
        public ShadingPatternValues? ShadingPattern {
            get {
                return _paragraphProperties?.Shading?.Val?.Value;
            }
            set {
                var props = _paragraph.ParagraphProperties ??= new ParagraphProperties();
                if (value != null) {
                    props.Shading ??= new Shading();
                    props.Shading.Val = value.Value;
                } else {
                    props.Shading?.Remove();
                }
            }
        }
    }
}
