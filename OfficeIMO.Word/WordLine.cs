using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a simple line shape inside a paragraph.
    /// </summary>
    public class WordLine : WordElement {
        internal WordDocument _document;
        internal WordParagraph _wordParagraph;
        internal Run _run;
        internal V.Line _line = null!;

        internal WordLine(WordDocument document, WordParagraph paragraph, double startXPt, double startYPt, double endXPt, double endYPt, string color = "#000000", double strokeWeightPt = 1) {
            _document = document;
            _wordParagraph = paragraph;
            var vmlStroke = color;
            if (!string.IsNullOrEmpty(vmlStroke) && !vmlStroke.StartsWith("#", StringComparison.Ordinal)) vmlStroke = "#" + vmlStroke;
            _line = new V.Line() {
                Id = "Line" + Guid.NewGuid().ToString("N"),
                Style = "mso-wrap-style:square",
                From = $"{startXPt}pt,{startYPt}pt",
                To = $"{endXPt}pt,{endYPt}pt",
                StrokeColor = vmlStroke,
                StrokeWeight = $"{strokeWeightPt}pt"
            };

            Picture pict = new Picture();
            pict.Append(_line);

            _run = paragraph.VerifyRun();
            _run.Append(pict);
        }

        internal WordLine(WordDocument document, Paragraph paragraph, Run run) {
            _document = document;
            _wordParagraph = new WordParagraph(document, paragraph, run);
            _run = run;
            _line = run.Descendants<V.Line>().FirstOrDefault()
                ?? throw new ArgumentException("The provided run does not contain a VML line.", nameof(run));
        }

        /// <summary>
        /// Gets or sets the stroke color as hexadecimal string.
        /// </summary>
        public string ColorHex {
            get {
                var v = _line.StrokeColor?.Value ?? string.Empty;
                if (v.StartsWith("#", StringComparison.Ordinal)) v = v.Substring(1);
                return v.ToLowerInvariant();
            }
            set {
                var v = value;
                if (!string.IsNullOrEmpty(v) && !v.StartsWith("#", StringComparison.Ordinal)) v = "#" + v;
                _line.StrokeColor = v;
            }
        }

        /// <summary>
        /// Gets or sets the stroke color using <see cref="SixLabors.ImageSharp.Color"/>.
        /// </summary>
        public Color Color {
            get {
                var color = _line.StrokeColor?.Value ?? "";
                if (!color.StartsWith("#", StringComparison.Ordinal)) color = "#" + color;
                return Color.Parse(color);
            }
            set {
                var hex = value.ToHexColor();
                if (!hex.StartsWith("#", StringComparison.Ordinal)) hex = "#" + hex;
                _line.StrokeColor = hex;
            }
        }

        /// <summary>
        /// Gets or sets the stroke weight in points.
        /// </summary>
        public double StrokeWeightPt {
            get {
                if (double.TryParse(_line.StrokeWeight?.Value?.Replace("pt", ""), out double value)) {
                    return value;
                }
                return 0;
            }
            set => _line.StrokeWeight = $"{value}pt";
        }

        /// <summary>
        /// Removes the line from the paragraph.
        /// </summary>
        public void Remove() {
            _run?.Remove();
        }
    }
}
