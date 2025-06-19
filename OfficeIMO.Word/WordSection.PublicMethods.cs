using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordSection {
        /// <summary>
        /// Updates the section margins using the specified <see cref="WordMargin"/>.
        /// </summary>
        public WordSection SetMargins(WordMargin pageMargins) {
            return WordMargins.SetMargins(this, pageMargins);
        }

        /// <summary>
        /// Adds a new paragraph to this section. When <paramref name="newRun"/> is
        /// <c>true</c> the paragraph starts with an empty run.
        /// </summary>
        public WordParagraph AddParagraph(bool newRun) {
            var wordParagraph = new WordParagraph(_document, newParagraph: true, newRun: newRun);
            if (this.Paragraphs.Count == 0) {
                WordParagraph paragraph = this._document.AddParagraph(wordParagraph);
                return paragraph;
            } else {
                WordParagraph lastParagraphWithinSection = this.Paragraphs.Last();
                WordParagraph paragraph = lastParagraphWithinSection.AddParagraphAfterSelf(this, wordParagraph);
                return paragraph;
            }
        }

        /// <summary>
        /// Adds a new paragraph containing the specified text to this section.
        /// </summary>
        public WordParagraph AddParagraph(string text = "") {
            if (this.Paragraphs.Count == 0) {
                WordParagraph paragraph = this._document.AddParagraph();
                if (text != "") {
                    paragraph.Text = text;
                }

                return paragraph;
            } else {
                WordParagraph lastParagraphWithinSection = this.Paragraphs.Last();

                WordParagraph paragraph = lastParagraphWithinSection.AddParagraphAfterSelf(this);
                paragraph._document = this._document;
                if (text != "") {
                    paragraph.Text = text;
                }

                return paragraph;
            }
        }

        /// <summary>
        /// Inserts a watermark into the section using either text or an image file.
        /// </summary>
        public WordWatermark AddWatermark(WordWatermarkStyle watermarkStyle, string textOrFilePath) {
            // return new WordWatermark(this._document, this, this.Header.Default, watermarkStyle, text);
            return new WordWatermark(this._document, this, watermarkStyle, textOrFilePath);
        }

        /// <summary>
        /// Applies borders around the page using the provided border settings.
        /// </summary>
        public WordSection SetBorders(WordBorder wordBorder) {
            this.Borders.SetBorder(wordBorder);

            return this;
        }

        /// <summary>
        /// Inserts a horizontal line formatted with the specified border style.
        /// </summary>
        public WordParagraph AddHorizontalLine(BorderValues? lineType = null, SixLabors.ImageSharp.Color? color = null, uint size = 12, uint space = 1) {
            lineType ??= BorderValues.Single;
            return this.AddParagraph("").AddHorizontalLine(lineType.Value, color, size, space);
        }

        /// <summary>
        /// Adds a hyperlink paragraph that navigates to the specified external URI.
        /// </summary>
        public WordParagraph AddHyperLink(string text, Uri uri, bool addStyle = false, string tooltip = "", bool history = true) {
            return this.AddParagraph("").AddHyperLink(text, uri, addStyle, tooltip, history);
        }

        /// <summary>
        /// Adds a hyperlink paragraph that points to an internal bookmark.
        /// </summary>
        public WordParagraph AddHyperLink(string text, string anchor, bool addStyle = false, string tooltip = "", bool history = true) {
            return this.AddParagraph("").AddHyperLink(text, anchor, addStyle, tooltip, history);
        }

        /// <summary>
        /// Creates default headers and footers for this section if they do not already exist.
        /// </summary>
        public void AddHeadersAndFooters() {
            WordHeadersAndFooters.AddHeadersAndFooters(this);
        }

        /// <summary>
        /// Appends a new paragraph containing a text box with the supplied text.
        /// </summary>
        public WordTextBox AddTextBox(string text, WrapTextImage wrapTextImage = WrapTextImage.Square) {
            return AddParagraph(newRun: true).AddTextBox(text, wrapTextImage);
        }

        /// <summary>
        /// Sets up default footnote properties such as numbering format and position for the section.
        /// </summary>
        public WordSection AddFootnoteProperties(NumberFormatValues? numberingFormat = null,
            FootnotePositionValues? position = null,
            RestartNumberValues? restartNumbering = null,
            int? startNumber = null) {
            var props = _sectionProperties.GetFirstChild<FootnoteProperties>();
            if (props == null) {
                props = new FootnoteProperties();
                _sectionProperties.Append(props);
            }

            props.RemoveAllChildren<NumberingFormat>();
            props.RemoveAllChildren<FootnotePosition>();
            props.RemoveAllChildren<NumberingRestart>();
            props.RemoveAllChildren<NumberingStart>();

            if (numberingFormat != null) {
                props.Append(new NumberingFormat() { Val = numberingFormat });
            }

            if (position != null) {
                props.Append(new FootnotePosition() { Val = position });
            }

            if (restartNumbering != null) {
                props.Append(new NumberingRestart() { Val = restartNumbering });
            }

            if (startNumber != null) {
                props.Append(new NumberingStart() { Val = (UInt16Value)startNumber.Value });
            }

            return this;
        }

        /// <summary>
        /// Sets up default endnote properties such as numbering format and position for the section.
        /// </summary>
        public WordSection AddEndnoteProperties(NumberFormatValues? numberingFormat = null,
            EndnotePositionValues? position = null,
            RestartNumberValues? restartNumbering = null,
            int? startNumber = null) {
            var props = _sectionProperties.GetFirstChild<EndnoteProperties>();
            if (props == null) {
                props = new EndnoteProperties();
                _sectionProperties.Append(props);
            }

            props.RemoveAllChildren<NumberingFormat>();
            props.RemoveAllChildren<EndnotePosition>();
            props.RemoveAllChildren<NumberingRestart>();
            props.RemoveAllChildren<NumberingStart>();

            if (numberingFormat != null) {
                props.Append(new NumberingFormat() { Val = numberingFormat });
            }

            if (position != null) {
                props.Append(new EndnotePosition() { Val = position });
            }

            if (restartNumbering != null) {
                props.Append(new NumberingRestart() { Val = restartNumbering });
            }

            if (startNumber != null) {
                props.Append(new NumberingStart() { Val = (UInt16Value)startNumber.Value });
            }

            return this;
        }

        /// <summary>
        /// Removes this section and all of its content from the document,
        /// cleaning up numbering and any unreferenced header and footer parts.
        /// </summary>
        public void RemoveSection() {
            foreach (var list in this.Lists.ToList()) {
                list.Remove();
            }

            foreach (var element in this.ElementsByType.ToList()) {
                switch (element) {
                    case WordParagraph paragraph:
                        paragraph.Remove();
                        break;
                    case WordTable table:
                        table.Remove();
                        break;
                    case WordTextBox textBox:
                        textBox.Remove();
                        break;
                    case WordImage image:
                        image.Remove();
                        break;
                    case WordEmbeddedDocument embedded:
                        embedded.Remove();
                        break;
                }
            }

            foreach (var headerRef in _sectionProperties.Elements<HeaderReference>().ToList()) {
                string id = headerRef.Id;
                bool usedElsewhere = _document.Sections
                    .Where(s => s != this)
                    .Any(s => s._sectionProperties.Elements<HeaderReference>().Any(hr => hr.Id == id));
                if (!usedElsewhere) {
                    var part = (HeaderPart)_document._wordprocessingDocument.MainDocumentPart.GetPartById(id);
                    _document._wordprocessingDocument.MainDocumentPart.DeletePart(part);
                }
                headerRef.Remove();
            }

            foreach (var footerRef in _sectionProperties.Elements<FooterReference>().ToList()) {
                string id = footerRef.Id;
                bool usedElsewhere = _document.Sections
                    .Where(s => s != this)
                    .Any(s => s._sectionProperties.Elements<FooterReference>().Any(fr => fr.Id == id));
                if (!usedElsewhere) {
                    var part = (FooterPart)_document._wordprocessingDocument.MainDocumentPart.GetPartById(id);
                    _document._wordprocessingDocument.MainDocumentPart.DeletePart(part);
                }
                footerRef.Remove();
            }

            if (_sectionProperties.Parent is Paragraph p) {
                p.Remove();
            } else if (_sectionProperties.Parent != null) {
                _sectionProperties.Remove();
            }

            _document.Sections.Remove(this);
        }
    }
}
