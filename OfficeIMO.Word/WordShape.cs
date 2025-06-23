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
        /// Identifier of the shape.
        /// </summary>
        public string Id {
            get {
                if (_rectangle != null) return _rectangle.Id?.Value ?? string.Empty;
                if (_ellipse != null) return _ellipse.Id?.Value ?? string.Empty;
                if (_polygon != null) return _polygon.Id?.Value ?? string.Empty;
                if (_line != null) return _line.Id?.Value ?? string.Empty;
                return string.Empty;
            }
            set {
                if (_rectangle != null) _rectangle.Id = value;
                if (_ellipse != null) _ellipse.Id = value;
                if (_polygon != null) _polygon.Id = value;
                if (_line != null) _line.Id = value;
            }
        }

        /// <summary>
        /// Optional title of the shape.
        /// </summary>
        public string? Title {
            get {
                if (_rectangle != null) return _rectangle.Title?.Value;
                if (_ellipse != null) return _ellipse.Title?.Value;
                if (_polygon != null) return _polygon.Title?.Value;
                if (_line != null) return _line.Title?.Value;
                return null;
            }
            set {
                if (_rectangle != null) _rectangle.Title = value;
                if (_ellipse != null) _ellipse.Title = value;
                if (_polygon != null) _polygon.Title = value;
                if (_line != null) _line.Title = value;
            }
        }

        /// <summary>
        /// Alternative text description of the shape.
        /// </summary>
        public string? Description {
            get {
                if (_rectangle != null) return _rectangle.Alternate?.Value;
                if (_ellipse != null) return _ellipse.Alternate?.Value;
                if (_polygon != null) return _polygon.Alternate?.Value;
                if (_line != null) return _line.Alternate?.Value;
                return null;
            }
            set {
                if (_rectangle != null) _rectangle.Alternate = value;
                if (_ellipse != null) _ellipse.Alternate = value;
                if (_polygon != null) _polygon.Alternate = value;
                if (_line != null) _line.Alternate = value;
            }
        }

        /// <summary>
        /// Whether the shape is hidden.
        /// </summary>
        public bool? Hidden {
            get {
                if (_rectangle?.Hidden != null) return _rectangle.Hidden.Value;
                if (_ellipse?.Hidden != null) return _ellipse.Hidden.Value;
                if (_polygon?.Hidden != null) return _polygon.Hidden.Value;
                if (_line?.Hidden != null) return _line.Hidden.Value;
                return null;
            }
            set {
                if (_rectangle != null) _rectangle.Hidden = value;
                if (_ellipse != null) _ellipse.Hidden = value;
                if (_polygon != null) _polygon.Hidden = value;
                if (_line != null) _line.Hidden = value;
            }
        }

        /// <summary>
        /// Outline color in hex format. Null when not applicable.
        /// </summary>
        public string? StrokeColorHex {
            get {
                if (_rectangle != null) return _rectangle.StrokeColor?.Value;
                if (_ellipse != null) return _ellipse.StrokeColor?.Value;
                if (_polygon != null) return _polygon.StrokeColor?.Value;
                if (_line != null) return _line.StrokeColor?.Value;
                return null;
            }
            set {
                if (_rectangle != null) _rectangle.StrokeColor = value;
                if (_ellipse != null) _ellipse.StrokeColor = value;
                if (_polygon != null) _polygon.StrokeColor = value;
                if (_line != null) _line.StrokeColor = value;
            }
        }

        /// <summary>
        /// Outline color using <see cref="SixLabors.ImageSharp.Color"/>.
        /// </summary>
        public SixLabors.ImageSharp.Color StrokeColor {
            get {
                var hex = StrokeColorHex;
                if (string.IsNullOrEmpty(hex)) return SixLabors.ImageSharp.Color.Transparent;
                if (!hex.StartsWith("#")) hex = "#" + hex;
                return SixLabors.ImageSharp.Color.Parse(hex);
            }
            set => StrokeColorHex = value.ToHexColor();
        }

        /// <summary>
        /// Outline thickness in points.
        /// </summary>
        public double? StrokeWeight {
            get {
                string? v = null;
                if (_rectangle != null) v = _rectangle.StrokeWeight?.Value;
                if (_ellipse != null) v ??= _ellipse.StrokeWeight?.Value;
                if (_polygon != null) v ??= _polygon.StrokeWeight?.Value;
                if (_line != null) v ??= _line.StrokeWeight?.Value;
                if (string.IsNullOrEmpty(v)) return null;
                return double.Parse(v.Replace("pt", string.Empty), CultureInfo.InvariantCulture);
            }
            set {
                string? v = value != null ? $"{value.Value.ToString(CultureInfo.InvariantCulture)}pt" : null;
                if (_rectangle != null) _rectangle.StrokeWeight = v;
                if (_ellipse != null) _ellipse.StrokeWeight = v;
                if (_polygon != null) _polygon.StrokeWeight = v;
                if (_line != null) _line.StrokeWeight = v;
            }
        }

        /// <summary>
        /// Determines whether the outline is drawn.
        /// </summary>
        public bool? Stroked {
            get {
                if (_rectangle?.Stroked != null) return _rectangle.Stroked.Value;
                if (_ellipse?.Stroked != null) return _ellipse.Stroked.Value;
                if (_polygon?.Stroked != null) return _polygon.Stroked.Value;
                if (_line?.Stroked != null) return _line.Stroked.Value;
                return null;
            }
            set {
                if (_rectangle != null) _rectangle.Stroked = value;
                if (_ellipse != null) _ellipse.Stroked = value;
                if (_polygon != null) _polygon.Stroked = value;
                if (_line != null) _line.Stroked = value;
            }
        }

        private string? GetStyle() {
            return _rectangle?.Style?.Value ?? _ellipse?.Style?.Value ?? _polygon?.Style?.Value;
        }

        private void SetStyle(string style) {
            if (_rectangle != null) _rectangle.Style = style;
            if (_ellipse != null) _ellipse.Style = style;
            if (_polygon != null) _polygon.Style = style;
        }

        private string? GetStyleValue(string name) {
            var style = GetStyle();
            if (string.IsNullOrEmpty(style)) return null;
            foreach (var part in style.Split(';')) {
                var kv = part.Split(':');
                if (kv.Length == 2 && kv[0] == name) return kv[1];
            }
            return null;
        }

        private void SetStyleValue(string name, string value) {
            var style = GetStyle() ?? string.Empty;
            var parts = style.Split(';', StringSplitOptions.RemoveEmptyEntries).ToList();
            bool updated = false;
            for (int i = 0; i < parts.Count; i++) {
                var kv = parts[i].Split(':');
                if (kv.Length == 2 && kv[0] == name) {
                    parts[i] = $"{name}:{value}";
                    updated = true;
                    break;
                }
            }
            if (!updated) parts.Add($"{name}:{value}");
            SetStyle(string.Join(";", parts));
        }

        private void RemoveStyleValue(string name) {
            var style = GetStyle();
            if (string.IsNullOrEmpty(style)) return;
            var parts = style.Split(';', StringSplitOptions.RemoveEmptyEntries).ToList();
            parts.RemoveAll(p => p.Split(':').FirstOrDefault() == name);
            SetStyle(string.Join(";", parts));
        }

        /// <summary>
        /// Width of the shape in points.
        /// </summary>
        public double Width {
            get {
                var v = GetStyleValue("width");
                if (!string.IsNullOrEmpty(v)) return double.Parse(v.Replace("pt", string.Empty), CultureInfo.InvariantCulture);
                return 0;
            }
            set => SetStyleValue("width", $"{value.ToString(CultureInfo.InvariantCulture)}pt");
        }

        /// <summary>
        /// Height of the shape in points.
        /// </summary>
        public double Height {
            get {
                var v = GetStyleValue("height");
                if (!string.IsNullOrEmpty(v)) return double.Parse(v.Replace("pt", string.Empty), CultureInfo.InvariantCulture);
                return 0;
            }
            set => SetStyleValue("height", $"{value.ToString(CultureInfo.InvariantCulture)}pt");
        }

        /// <summary>
        /// Left position of the shape in points. Returns null when not explicitly set.
        /// </summary>
        public double? Left {
            get {
                var v = GetStyleValue("margin-left");
                if (string.IsNullOrEmpty(v)) return null;
                return double.Parse(v.Replace("pt", string.Empty), CultureInfo.InvariantCulture);
            }
            set {
                if (value == null) {
                    RemoveStyleValue("margin-left");
                } else {
                    SetStyleValue("margin-left", $"{value.Value.ToString(CultureInfo.InvariantCulture)}pt");
                    if (GetStyleValue("position") == null) SetStyleValue("position", "absolute");
                }
            }
        }

        /// <summary>
        /// Top position of the shape in points. Returns null when not explicitly set.
        /// </summary>
        public double? Top {
            get {
                var v = GetStyleValue("margin-top");
                if (string.IsNullOrEmpty(v)) return null;
                return double.Parse(v.Replace("pt", string.Empty), CultureInfo.InvariantCulture);
            }
            set {
                if (value == null) {
                    RemoveStyleValue("margin-top");
                } else {
                    SetStyleValue("margin-top", $"{value.Value.ToString(CultureInfo.InvariantCulture)}pt");
                    if (GetStyleValue("position") == null) SetStyleValue("position", "absolute");
                }
            }
        }

        /// <summary>
        /// Rotation of the shape in degrees. Returns null when not set.
        /// </summary>
        public double? Rotation {
            get {
                var v = GetStyleValue("rotation");
                if (string.IsNullOrEmpty(v)) return null;
                return double.Parse(v, CultureInfo.InvariantCulture);
            }
            set {
                if (value == null) {
                    RemoveStyleValue("rotation");
                } else {
                    SetStyleValue("rotation", value.Value.ToString(CultureInfo.InvariantCulture));
                }
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
