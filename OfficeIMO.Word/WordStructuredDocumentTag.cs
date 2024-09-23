using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeIMO.Word {
    public class WordStructuredDocumentTag : WordElement {
        private WordDocument _document;
        private Paragraph _paragraph;
        private SdtRun _stdRun;

        public string Alias {
            get {
                var sdtAlias = _stdRun.SdtProperties.OfType<SdtAlias>().FirstOrDefault();
                if (sdtAlias != null) {
                    return sdtAlias.Val;
                }

                return null;
            }
        }

        public string Text {
            get {
                var run = _stdRun.SdtContentRun.ChildElements.OfType<Run>().FirstOrDefault();
                if (run != null) {
                    var text = run.OfType<Text>().FirstOrDefault();
                    if (text != null) {
                        return text.Text;
                    }
                }
                return null;
            }
            set {
                var run = _stdRun.SdtContentRun.ChildElements.OfType<Run>().FirstOrDefault();
                if (run != null) {
                    var text = run.OfType<Text>().FirstOrDefault();
                    if (text != null) {
                        text.Text = value;
                    }
                }
            }
        }

        public WordStructuredDocumentTag(WordDocument document, Paragraph paragraph, SdtRun stdRun) {
            this._document = document;
            this._paragraph = paragraph;
            this._stdRun = stdRun;
        }

        public void Remove() {
            if (this._stdRun != null) {
                this._stdRun.Remove();
            }
        }
    }
}
