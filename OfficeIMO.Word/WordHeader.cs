using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class WordHeaders {
        public WordHeader Default {
            get;
            set;
        }
        public WordHeader Even {
            get;
            set;
        }
        public WordHeader First {
            get;
            set;
        }
    }
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

        public WordWatermark AddWatermark(WordWatermarkStyle watermarkStyle, string textOrFilePath) {
            return new WordWatermark(this._document, this._section, this, watermarkStyle, textOrFilePath);
        }

        public WordTextBox AddTextBox(string text, WrapTextImage wrapTextImage = WrapTextImage.Square) {
            WordTextBox wordTextBox = new WordTextBox(this._document, this, text, wrapTextImage);
            return wordTextBox;
        }
    }
}
