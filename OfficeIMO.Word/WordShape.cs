using System;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents simple rectangle shape inside a paragraph.
    /// </summary>
    public class WordShape : WordElement {
        internal WordDocument _document;
        internal WordParagraph _wordParagraph;
        internal Run _run;
        internal V.Rectangle _rectangle;

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

            _run = new Run();
            _run.Append(pict);
            paragraph._paragraph.Append(_run);
        }

        internal WordShape(WordDocument document, Paragraph paragraph, Run run) {
            _document = document;
            _wordParagraph = new WordParagraph(document, paragraph, run);
            _run = run;
            _rectangle = run.Descendants<V.Rectangle>().FirstOrDefault();
        }

        /// <summary>
        /// Width of the shape in points.
        /// </summary>
        public double Width {
            get {
            var style = _rectangle.Style?.Value;
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
            var style = _rectangle.Style?.Value;
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
