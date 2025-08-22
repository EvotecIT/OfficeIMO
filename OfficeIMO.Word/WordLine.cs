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
        internal V.Line _line;

        internal WordLine(WordDocument document, WordParagraph paragraph, double startXPt, double startYPt, double endXPt, double endYPt, string color = "#000000", double strokeWeightPt = 1) {
            _document = document;
            _wordParagraph = paragraph;
            _line = new V.Line() {
                Id = "Line" + Guid.NewGuid().ToString("N"),
                Style = "mso-wrap-style:square",
                From = $"{startXPt}pt,{startYPt}pt",
                To = $"{endXPt}pt,{endYPt}pt",
                StrokeColor = color,
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
                ?? throw new InvalidOperationException("Line element not found in run.");
        }

        /// <summary>
        /// Gets or sets the stroke color as hexadecimal string.
        /// </summary>
        public string ColorHex {
            get => _line.StrokeColor?.Value ?? string.Empty;
            set => _line.StrokeColor = value;
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
            set => _line.StrokeColor = value.ToHexColor();
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
