using DocumentFormat.OpenXml.Wordprocessing;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>Native paragraph shading fill (hexadecimal or <c>auto</c>); empty removes the shading.</summary>
        internal string ShadingFillValue {
            get {
                var fill = _paragraphProperties?.Shading?.Fill?.Value;
                return fill != null ? fill.ToUpperInvariant() : string.Empty;
            }
            set {
                var props = _paragraph.ParagraphProperties ??= new ParagraphProperties();
                if (value != string.Empty) {
                    var color = value.Replace("#", string.Empty).ToUpperInvariant();
                    props.Shading ??= new Shading();
                    props.Shading.Fill = color;
                    props.Shading.Val ??= ShadingPatternValues.Clear;
                } else {
                    props.Shading?.Remove();
                }
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
                return _paragraphProperties?.Shading?.Val?.Value.ToOfficeEnum();
            }
            set {
                var props = _paragraph.ParagraphProperties ??= new ParagraphProperties();
                if (value != null) {
                    props.Shading ??= new Shading();
                    props.Shading.Val = value.Value.ToOpenXml();
                } else {
                    props.Shading?.Remove();
                }
            }
        }
    }
}
