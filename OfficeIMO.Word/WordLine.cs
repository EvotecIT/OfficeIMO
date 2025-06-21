using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;

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
            _line = run.Descendants<V.Line>().FirstOrDefault();
        }

        /// <summary>
        /// Removes the line from the paragraph.
        /// </summary>
        public void Remove() {
            _run?.Remove();
        }
    }
}
