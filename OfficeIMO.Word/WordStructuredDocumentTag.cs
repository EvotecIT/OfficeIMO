using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a structured document tag (content control) within a paragraph.
    /// </summary>
    public class WordStructuredDocumentTag : WordElement {
        private WordDocument _document;
        private Paragraph _paragraph;
        private SdtRun _stdRun;

        /// <summary>
        /// Gets the alias of this structured document tag.
        /// </summary>
        public string Alias {
            get {
                var sdtAlias = _stdRun.SdtProperties.OfType<SdtAlias>().FirstOrDefault();
                if (sdtAlias != null) {
                    return sdtAlias.Val;
                }

                return null;
            }
        }

        /// <summary>
        /// Gets or sets the tag associated with this structured document tag.
        /// </summary>
        public string Tag {
            get {
                var tag = _stdRun.SdtProperties.OfType<Tag>().FirstOrDefault();
                return tag?.Val;
            }
            set {
                var tag = _stdRun.SdtProperties.OfType<Tag>().FirstOrDefault();
                if (tag == null) {
                    tag = new Tag();
                    _stdRun.SdtProperties.Append(tag);
                }
                tag.Val = value;
            }
        }

        /// <summary>
        /// Gets or sets the text content of this structured document tag.
        /// </summary>
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

        /// <summary>
        /// Initializes a new instance of the <see cref="WordStructuredDocumentTag"/> class.
        /// </summary>
        /// <param name="document">Parent document.</param>
        /// <param name="paragraph">Paragraph containing the tag.</param>
        /// <param name="stdRun">Underlying structured document run.</param>
        public WordStructuredDocumentTag(WordDocument document, Paragraph paragraph, SdtRun stdRun) {
            this._document = document;
            this._paragraph = paragraph;
            this._stdRun = stdRun;
        }

        /// <summary>
        /// Removes this structured document tag from the document.
        /// </summary>
        public void Remove() {
            if (this._stdRun != null) {
                this._stdRun.Remove();
            }
        }
    }
}
