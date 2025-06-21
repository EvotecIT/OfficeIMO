using System;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents simple VML shapes inside a paragraph.
    /// </summary>
    public class WordShape : WordElement {
        /// <summary>Parent document.</summary>
        internal WordDocument _document;
        /// <summary>Parent paragraph.</summary>
        internal WordParagraph _wordParagraph;
        /// <summary>Run that hosts the shape.</summary>
        internal Run _run;
        /// <summary>The rectangle element if present.</summary>
        internal V.Rectangle _rectangle;
        /// <summary>The ellipse element if present.</summary>
        internal V.Oval _ellipse;
        /// <summary>The line element if present.</summary>
        internal V.Line _line;
        /// <summary>The polygon element if present.</summary>
        internal V.PolyLine _polygon;

        /// <summary>
        /// Initializes a new rectangle shape and appends it to the paragraph.
        /// </summary>
        internal WordShape(WordDocument document, WordParagraph paragraph, double widthPt, double heightPt, string fillColor = "#FFFFFF") {
            _document = document;
            _wordParagraph = paragraph;

            _rectangle = new V.Rectangle() {
                Id = "Rectangle" + Guid.NewGuid().ToString("N"),
                Style = $"width:{widthPt}pt;height:{heightPt}pt;mso-wrap-style:square",
                FillColor = fillColor,
                Stroked = false
            };

            Picture pict = new Picture();
            pict.Append(_rectangle);

            _run = paragraph.VerifyRun();
            _run.Append(pict);
        }

        /// <summary>
        /// Initializes a <see cref="WordShape"/> from existing run content.
        /// </summary>
        internal WordShape(WordDocument document, Paragraph paragraph, Run run) {
            _document = document;
            _wordParagraph = new WordParagraph(document, paragraph, run);
            _run = run;
            _rectangle = run.Descendants<V.Rectangle>().FirstOrDefault();
            _ellipse = run.Descendants<V.Oval>().FirstOrDefault();
            _line = run.Descendants<V.Line>().FirstOrDefault();
            _polygon = run.Descendants<V.PolyLine>().FirstOrDefault();
        }

        /// <summary>
        /// Adds an ellipse shape to the given paragraph.
        /// </summary>
        public static WordShape AddEllipse(WordParagraph paragraph, double widthPt, double heightPt, string fillColor = "#FFFFFF") {
            var ellipse = new V.Oval() {
                Id = "Ellipse" + Guid.NewGuid().ToString("N"),
                Style = $"width:{widthPt}pt;height:{heightPt}pt;mso-wrap-style:square",
                FillColor = fillColor,
                Stroked = false
            };

            Picture pict = new Picture();
            pict.Append(ellipse);

            var run = paragraph.VerifyRun();
            run.Append(pict);

            return new WordShape(paragraph._document, paragraph._paragraph, run);
        }

        /// <summary>
        /// Adds an ellipse shape using <see cref="SixLabors.ImageSharp.Color"/>.
        /// </summary>
        public static WordShape AddEllipse(WordParagraph paragraph, double widthPt, double heightPt, SixLabors.ImageSharp.Color fillColor) {
            return AddEllipse(paragraph, widthPt, heightPt, fillColor.ToHexColor());
        }

        /// <summary>
        /// Adds a line shape to the given paragraph.
        /// </summary>
        public static WordShape AddLine(WordParagraph paragraph, double startXPt, double startYPt, double endXPt, double endYPt, string color = "#000000", double strokeWeightPt = 1) {
            var line = new V.Line() {
                Id = "Line" + Guid.NewGuid().ToString("N"),
                Style = "mso-wrap-style:square",
                From = $"{startXPt}pt,{startYPt}pt",
                To = $"{endXPt}pt,{endYPt}pt",
                StrokeColor = color,
                StrokeWeight = $"{strokeWeightPt}pt"
            };

            Picture pict = new Picture();
            pict.Append(line);

            var run = paragraph.VerifyRun();
            run.Append(pict);

            return new WordShape(paragraph._document, paragraph._paragraph, run);
        }

        /// <summary>
        /// Adds a line shape using <see cref="SixLabors.ImageSharp.Color"/>.
        /// </summary>
        public static WordShape AddLine(WordParagraph paragraph, double startXPt, double startYPt, double endXPt, double endYPt, SixLabors.ImageSharp.Color color, double strokeWeightPt = 1) {
            return AddLine(paragraph, startXPt, startYPt, endXPt, endYPt, color.ToHexColor(), strokeWeightPt);
        }

        /// <summary>
        /// Adds a polygon shape to the given paragraph.
        /// </summary>
        public static WordShape AddPolygon(WordParagraph paragraph, string points, string fillColor = "#FFFFFF", string strokeColor = "#000000") {
            var poly = new V.PolyLine() {
                Id = "Polygon" + Guid.NewGuid().ToString("N"),
                Style = "mso-wrap-style:square",
                Points = points,
                FillColor = fillColor,
                StrokeColor = strokeColor
            };

            Picture pict = new Picture();
            pict.Append(poly);

            var run = paragraph.VerifyRun();
            run.Append(pict);

            return new WordShape(paragraph._document, paragraph._paragraph, run);
        }

        /// <summary>
        /// Adds a polygon shape using <see cref="SixLabors.ImageSharp.Color"/> values.
        /// </summary>
        public static WordShape AddPolygon(WordParagraph paragraph, string points, SixLabors.ImageSharp.Color fillColor, SixLabors.ImageSharp.Color strokeColor) {
            return AddPolygon(paragraph, points, fillColor.ToHexColor(), strokeColor.ToHexColor());
        }

        /// <summary>
        /// Gets or sets the fill color as hexadecimal string.
        /// </summary>
        public string FillColorHex {
            get {
                if (_rectangle != null) return _rectangle.FillColor?.Value;
                if (_ellipse != null) return _ellipse.FillColor?.Value;
                if (_polygon != null) return _polygon.FillColor?.Value;
                return string.Empty;
            }
            set {
                if (_rectangle != null) _rectangle.FillColor = value;
                if (_ellipse != null) _ellipse.FillColor = value;
                if (_polygon != null) _polygon.FillColor = value;
            }
        }

        /// <summary>
        /// Gets or sets the fill color using <see cref="SixLabors.ImageSharp.Color"/>.
        /// </summary>
        public SixLabors.ImageSharp.Color FillColor {
            get {
                var hex = FillColorHex;
                if (string.IsNullOrEmpty(hex)) return SixLabors.ImageSharp.Color.Transparent;
                if (!hex.StartsWith("#")) hex = "#" + hex;
                return SixLabors.ImageSharp.Color.Parse(hex);
            }
            set => FillColorHex = value.ToHexColor();
        }

        /// <summary>
        /// Width of the shape in points.
        /// </summary>
        public double Width {
            get {
                var style = _rectangle?.Style?.Value ?? _ellipse?.Style?.Value ?? _polygon?.Style?.Value;
                if (style != null) {
                    foreach (var part in style.Split(';')) {
                        var kv = part.Split(':');
                        if (kv.Length == 2 && kv[0] == "width") {
                            return double.Parse(kv[1].Replace("pt", ""), CultureInfo.InvariantCulture);
                        }
                    }
                }
                return 0;
            }
        }

        /// <summary>
        /// Height of the shape in points.
        /// </summary>
        public double Height {
            get {
                var style = _rectangle?.Style?.Value ?? _ellipse?.Style?.Value ?? _polygon?.Style?.Value;
                if (style != null) {
                    foreach (var part in style.Split(';')) {
                        var kv = part.Split(':');
                        if (kv.Length == 2 && kv[0] == "height") {
                            return double.Parse(kv[1].Replace("pt", ""), CultureInfo.InvariantCulture);
                        }
                    }
                }
                return 0;
            }
        }

        /// <summary>
        /// Removes the shape from the paragraph.
        /// </summary>
        public void Remove() {
            _run?.Remove();
        }
    }
}
