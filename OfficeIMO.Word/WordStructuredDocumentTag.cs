using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a structured document tag (content control) within a Word document.
    /// </summary>
    public class WordStructuredDocumentTag : WordElement {
        private WordDocument _document;
        private Paragraph? _paragraph;
        private SdtRun? _stdRun;
        private SdtBlock? _sdtBlock;

        /// <summary>
        /// Gets the alias associated with this content control.
        /// </summary>
        public string? Alias {
            get {
                if (_stdRun != null) {
                    var sdtAlias = _stdRun.SdtProperties?.OfType<SdtAlias>().FirstOrDefault();
                    return sdtAlias?.Val;
                }

                if (_sdtBlock != null) {
                    var sdtAlias = _sdtBlock.SdtProperties?.OfType<SdtAlias>().FirstOrDefault();
                    return sdtAlias?.Val;
                }

                return null;
            }
        }

        /// <summary>
        /// Gets or sets the tag value for this content control.
        /// </summary>
        public string? Tag {
            get {
                if (_stdRun != null) {
                    var tag = _stdRun.SdtProperties?.OfType<Tag>().FirstOrDefault();
                    return tag?.Val;
                }

                if (_sdtBlock != null) {
                    var tag = _sdtBlock.SdtProperties?.OfType<Tag>().FirstOrDefault();
                    return tag?.Val;
                }

                return null;
            }
            set {
                if (_stdRun != null) {
                    var properties = _stdRun.SdtProperties ??= new SdtProperties();
                    var tag = properties.OfType<Tag>().FirstOrDefault();
                    if (tag == null) {
                        tag = new Tag();
                        properties.Append(tag);
                    }
                    tag.Val = value;
                } else if (_sdtBlock != null) {
                    var properties = _sdtBlock.SdtProperties ??= new SdtProperties();
                    var tag = properties.OfType<Tag>().FirstOrDefault();
                    if (tag == null) {
                        tag = new Tag();
                        properties.Append(tag);
                    }
                    tag.Val = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets the inner text of the content control.
        /// </summary>
        public string? Text {
            get {
                if (_stdRun != null) {
                    var run = _stdRun.SdtContentRun?.ChildElements.OfType<Run>().FirstOrDefault();
                    var text = run?.OfType<Text>().FirstOrDefault();
                    return text?.Text;
                }

                if (_sdtBlock != null) {
                    var paragraph = _sdtBlock.SdtContentBlock?.ChildElements.OfType<Paragraph>().FirstOrDefault();
                    var run = paragraph?.ChildElements.OfType<Run>().FirstOrDefault();
                    var text = run?.OfType<Text>().FirstOrDefault();
                    return text?.Text;
                }

                return null;
            }
            set {
                if (_stdRun != null) {
                    var run = _stdRun.SdtContentRun?.ChildElements.OfType<Run>().FirstOrDefault();
                    if (run == null) {
                        run = new Run();
                        _stdRun.SdtContentRun ??= new SdtContentRun();
                        _stdRun.SdtContentRun.Append(run);
                    }

                    var text = run.OfType<Text>().FirstOrDefault();
                    if (text == null) {
                        text = new Text { Space = SpaceProcessingModeValues.Preserve };
                        run.Append(text);
                    }

                    text.Text = value ?? string.Empty;
                } else if (_sdtBlock != null) {
                    _sdtBlock.SdtContentBlock ??= new SdtContentBlock();
                    var paragraph = _sdtBlock.SdtContentBlock.ChildElements.OfType<Paragraph>().FirstOrDefault();
                    if (paragraph == null) {
                        paragraph = new Paragraph();
                        _sdtBlock.SdtContentBlock.Append(paragraph);
                    }

                    var run = paragraph.ChildElements.OfType<Run>().FirstOrDefault();
                    if (run == null) {
                        run = new Run();
                        paragraph.Append(run);
                    }

                    var text = run.OfType<Text>().FirstOrDefault();
                    if (text == null) {
                        text = new Text { Space = SpaceProcessingModeValues.Preserve };
                        run.Append(text);
                    }

                    text.Text = value ?? string.Empty;
                }
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordStructuredDocumentTag"/> class.
        /// </summary>
        /// <param name="document">Parent document.</param>
        /// <param name="paragraph">Paragraph that contains the content control.</param>
        /// <param name="stdRun">Underlying structured document run.</param>
        public WordStructuredDocumentTag(WordDocument document, Paragraph paragraph, SdtRun stdRun) {
            this._document = document;
            this._paragraph = paragraph;
            this._stdRun = stdRun;
        }

        /// <summary>
        /// Initializes a new instance from a structured document block.
        /// </summary>
        /// <param name="document">Parent document.</param>
        /// <param name="sdtBlock">Structured document block.</param>
        public WordStructuredDocumentTag(WordDocument document, SdtBlock sdtBlock) {
            this._document = document;
            this._sdtBlock = sdtBlock;
        }

        /// <summary>
        /// Removes the structured document tag from the document.
        /// </summary>
        public void Remove() {
            if (this._stdRun != null) {
                this._stdRun.Remove();
            } else if (this._sdtBlock != null) {
                this._sdtBlock.Remove();
            }
        }
    }
}
