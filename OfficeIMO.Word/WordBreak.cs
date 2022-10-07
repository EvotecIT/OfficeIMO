using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeIMO.Word {
    public class WordBreak {
        private WordDocument _document;
        private Paragraph _paragraph;
        private Run _run;

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

        public WordBreak(WordDocument document, Paragraph paragraph, Run run) {
            this._document = document;
            this._paragraph = paragraph;
            this._run = run;
        }

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
