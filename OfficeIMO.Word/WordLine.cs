using DocumentFormat.OpenXml.Wordprocessing;
using Color = OfficeIMO.OfficeColor;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a simple line shape inside a paragraph.
    /// </summary>
    public class WordLine : WordElement {
        internal WordDocument _document;
        internal WordParagraph _wordParagraph;
        internal Run _run;
        internal V.Line _line = null!;

        internal WordLine(WordDocument document, WordParagraph paragraph, OfficeLength startX, OfficeLength startY, OfficeLength endX, OfficeLength endY, string color, OfficeLength strokeWeight) {
            _document = document;
            _wordParagraph = paragraph;
            var vmlStroke = color;
            if (!string.IsNullOrEmpty(vmlStroke) && !vmlStroke.StartsWith("#", StringComparison.Ordinal)) vmlStroke = "#" + vmlStroke;
            _line = new V.Line() {
                Id = "Line" + Guid.NewGuid().ToString("N"),
                Style = "mso-wrap-style:square",
                From = $"{WordShape.FormatVmlLength(startX)},{WordShape.FormatVmlLength(startY)}",
                To = $"{WordShape.FormatVmlLength(endX)},{WordShape.FormatVmlLength(endY)}",
                StrokeColor = vmlStroke,
                StrokeWeight = WordShape.FormatVmlLength(strokeWeight)
            };

            Picture pict = new Picture();
            pict.Append(_line);

            _run = paragraph.VerifyRun();
            _run.Append(pict);
        }

        internal WordLine(WordDocument document, Paragraph paragraph, Run run, V.Line? selectedLine = null) {
            _document = document;
            _wordParagraph = new WordParagraph(document, paragraph, run);
            _run = run;
            _line = selectedLine ?? run.Descendants<V.Line>().FirstOrDefault()
                ?? throw new ArgumentException("The provided run does not contain a VML line.", nameof(run));
        }

        /// <summary>
        /// Gets or sets the stroke color as hexadecimal string.
        /// </summary>
        public string ColorHex {
            get {
                var v = _line.StrokeColor?.Value ?? string.Empty;
                if (v.StartsWith("#", StringComparison.Ordinal)) v = v.Substring(1);
                return v.ToUpperInvariant();
            }
            set {
                var v = value;
                if (!string.IsNullOrEmpty(v) && !v.StartsWith("#", StringComparison.Ordinal)) v = "#" + v;
                _line.StrokeColor = v;
            }
        }

        /// <summary>
        /// Gets or sets the stroke color using <see cref="OfficeIMO.OfficeColor"/>.
        /// </summary>
        public Color Color {
            get {
                var color = _line.StrokeColor?.Value;
                return Color.TryParse(color, out Color parsed) ? parsed : Color.Black;
            }
            set {
                var hex = value.ToRgbHex();
                if (!hex.StartsWith("#", StringComparison.Ordinal)) hex = "#" + hex;
                _line.StrokeColor = hex;
            }
        }

        /// <summary>
        /// Gets or sets the stroke weight in points.
        /// </summary>
        internal double StrokeWeightPoints {
            get {
                if (double.TryParse(_line.StrokeWeight?.Value?.Replace("pt", ""), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double value)) {
                    return value;
                }
                return 0;
            }
            set => _line.StrokeWeight = FormattableString.Invariant($"{value}pt");
        }

        /// <summary>Gets or sets the physical stroke thickness.</summary>
        public OfficeLength StrokeWeight {
            get => OfficeLength.TryParse(_line.StrokeWeight?.Value, out var value) ? value : OfficeLength.Zero;
            set => _line.StrokeWeight = WordShape.FormatVmlLength(value);
        }

        /// <summary>
        /// Removes the line from the paragraph.
        /// </summary>
        public void Remove() {
            _run?.Remove();
        }
    }
}
