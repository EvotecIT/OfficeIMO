using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a break in the text.
    /// Be it page break, soft break, column or text wrapping
    /// </summary>
    public class WordBreak {
        private WordDocument _document;
        private readonly Paragraph _paragraph;
        private readonly Run _run;

        /// <summary>
        /// Get type of Break in given paragraph
        /// </summary>
        public BreakValues? BreakType {
            get {
                if (_run != null) {
                    var brake = _run.ChildElements.OfType<Break>().FirstOrDefault();
                    if (brake == null) {
                        return null;
                    }

                    return brake.Type;
                }

                return null;
            }
        }

        /// <summary>
        /// Create new instance of WordBreak
        /// </summary>
        /// <param name="document"></param>
        /// <param name="paragraph"></param>
        /// <param name="run"></param>
        public WordBreak(WordDocument document, Paragraph paragraph, Run run) {
            this._document = document;
            this._paragraph = paragraph;
            this._run = run;
        }

        /// <summary>
        /// Remove the break from WordDocument. By default it removes break without removing paragraph.
        /// If you want paragraph removed please use IncludingParagraph bool.
        /// Please remember a paragraph can hold multiple other elements.
        /// </summary>
        /// <param name="includingParagraph"></param>
        public void Remove(bool includingParagraph = false) {
            if (includingParagraph) {
                this._paragraph.Remove();
            } else {
                if (_run.ChildElements.Count == 1) {
                    this._run.Remove();
                } else {
                    this._run.ChildElements.OfType<Break>().FirstOrDefault()?.Remove();
                }
            }
        }
    }
}
