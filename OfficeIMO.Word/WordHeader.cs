using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents the collection of headers used within a section.
    /// </summary>
    public class WordHeaders {
        /// <summary>
        /// Gets or sets the default header.
        /// </summary>
        public WordHeader Default { get; set; }

        /// <summary>
        /// Gets or sets the header for even pages.
        /// </summary>
        public WordHeader Even { get; set; }

        /// <summary>
        /// Gets or sets the header for the first page.
        /// </summary>
        public WordHeader First { get; set; }
    }
    /// <summary>
    /// Represents a single header instance within a section.
    /// </summary>
    public partial class WordHeader : WordHeaderFooter {
        private readonly WordSection _section;

        internal WordHeader(WordDocument document, HeaderReference headerReference, WordSection section) {
            _document = document;
            _id = headerReference.Id;
            _type = WordSection.GetType(headerReference.Type);
            _section = section;

            var listHeaders = document._wordprocessingDocument.MainDocumentPart.HeaderParts.ToList();
            foreach (HeaderPart headerPart in listHeaders) {
                var id = document._wordprocessingDocument.MainDocumentPart.GetIdOfPart(headerPart);
                if (id == _id) {
                    _headerPart = headerPart;
                    _header = headerPart.Header;
                }
            }

            if (_type == HeaderFooterValues.Default) {
                document._currentSection.Header.Default = this;
            } else if (_type == HeaderFooterValues.Even) {
                document._currentSection.Header.Even = this;
            } else if (_type == HeaderFooterValues.First) {
                document._currentSection.Header.First = this;
            } else {
                throw new InvalidOperationException("Shouldn't happen?");
            }
        }
        internal WordHeader(WordDocument document, HeaderFooterValues type, Header headerPartHeader, WordSection section) {
            _document = document;
            _header = headerPartHeader;
            _type = type;
            _section = section;
        }

        /// <summary>
        /// Adds a page number field to the header.
        /// </summary>
        /// <param name="wordPageNumberStyle">The numbering style to apply.</param>
        /// <returns>The created <see cref="WordPageNumber"/>.</returns>
        public WordPageNumber AddPageNumber(WordPageNumberStyle wordPageNumberStyle) {
            var pageNumber = new WordPageNumber(_document, this, wordPageNumberStyle);
            return pageNumber;
        }

        /// <summary>
        /// Removes headers from the provided <see cref="WordprocessingDocument"/>.
        /// When no <paramref name="types"/> are specified all headers are removed.
        /// </summary>
        /// <param name="wordprocessingDocument">Document to operate on.</param>
        /// <param name="types">Header types to remove.</param>
        public static void RemoveHeaders(WordprocessingDocument wordprocessingDocument, params HeaderFooterValues[] types) {
            var docPart = wordprocessingDocument.MainDocumentPart;
            DocumentFormat.OpenXml.Wordprocessing.Document document = docPart.Document;

            if (types == null || types.Length == 0) {
                if (docPart.HeaderParts.Any()) {
                    docPart.DeleteParts(docPart.HeaderParts);
                    var headers = document.Descendants<HeaderReference>().ToList();
                    foreach (var header in headers) {
                        header.Remove();
                    }
                }
                return;
            }

            var partsToDelete = new HashSet<HeaderPart>();
            var headersToRemove = document.Descendants<HeaderReference>()
                .Where(h => types.Contains(h.Type)).ToList();
            foreach (var header in headersToRemove) {
                var part = docPart.GetPartById(header.Id) as HeaderPart;
                if (part != null) {
                    partsToDelete.Add(part);
                }
                header.Remove();
            }
            foreach (var part in partsToDelete) {
                docPart.DeletePart(part);
            }
        }
        /// <summary>
        /// Removes headers from the specified <see cref="WordDocument"/>.
        /// When no <paramref name="types"/> are provided all headers are removed.
        /// </summary>
        /// <param name="document">Document to operate on.</param>
        /// <param name="types">Header types to remove.</param>
        public static void RemoveHeaders(WordDocument document, params HeaderFooterValues[] types) {
            RemoveHeaders(document._wordprocessingDocument, types);
        }

        /// <summary>
        /// Adds a watermark to the header.
        /// </summary>
        /// <param name="watermarkStyle">Watermark style.</param>
        /// <param name="textOrFilePath">Text or image path for the watermark.</param>
        /// <param name="horizontalOffset">Horizontal offset in points.</param>
        /// <param name="verticalOffset">Vertical offset in points.</param>
        /// <param name="scale">Scale factor for width and height.</param>
        /// <returns>The created <see cref="WordWatermark"/>.</returns>
        public WordWatermark AddWatermark(WordWatermarkStyle watermarkStyle, string textOrFilePath, double? horizontalOffset = null, double? verticalOffset = null, double scale = 1.0) {
            return new WordWatermark(this._document, this._section, this, watermarkStyle, textOrFilePath, horizontalOffset, verticalOffset, scale);
        }

        /// <summary>
        /// Adds a text box to the header.
        /// </summary>
        /// <param name="text">Text contained in the text box.</param>
        /// <param name="wrapTextImage">Wrapping style.</param>
        /// <returns>The created <see cref="WordTextBox"/>.</returns>
        public WordTextBox AddTextBox(string text, WrapTextImage wrapTextImage = WrapTextImage.Square) {
            WordTextBox wordTextBox = new WordTextBox(this._document, this, text, wrapTextImage);
            return wordTextBox;
        }

        /// <summary>
        /// Adds a VML text box to the header.
        /// </summary>
        public WordTextBox AddTextBoxVml(string text) {
            var paragraph = AddParagraph(newRun: true);
            return paragraph.AddTextBoxVml(text);
        }

        /// <summary>
        /// Adds a VML image to the header.
        /// </summary>
        public WordImage AddImageVml(string filePathImage, double? width = null, double? height = null) {
            var paragraph = AddParagraph(newRun: true);
            paragraph.AddImageVml(filePathImage, width, height);
            return paragraph.Image;
        }
    }
}
