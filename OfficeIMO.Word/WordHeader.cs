using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class WordHeader {
        public List<WordParagraph> Paragraphs {
            get {
                List<WordParagraph> paragraphs = new List<WordParagraph>();
                if (_header != null) {
                    var list = _header.ChildElements.OfType<Paragraph>();
                    foreach (var paragraph in list) {
                        paragraphs.Add(new WordParagraph(_document, paragraph, null));
                    }
                }

                return paragraphs;
            }
        }
        private readonly HeaderFooterValues _type;
        private readonly HeaderPart _headerPart;
        internal readonly Header _header;
        private string _id;
        private WordDocument _document;

        internal WordHeader(WordDocument document, HeaderReference headerReference) {
            _document = document;
            _id = headerReference.Id;
            _type = WordSection.GetType(headerReference.Type);

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
        internal WordHeader(WordDocument document, HeaderFooterValues type, Header headerPartHeader) {
            _document = document;
            _header = headerPartHeader;
            //if (type == HeaderFooterValues.First) {
            //    _headerFirst = headerPartHeader;
            //} else if (type == HeaderFooterValues.Default) {
            //    _headerDefault = headerPartHeader;
            //} else if (type == HeaderFooterValues.Even) {
            //    _headerEven = headerPartHeader;
            //}
            _type = type;
        }
        public WordParagraph AddParagraph() {
            var wordParagraph = new WordParagraph();
            //if (_type == HeaderFooterValues.First) {
            //    _headerFirst.Append(wordParagraph._paragraph);
            //} else if (_type == HeaderFooterValues.Default) {
            //    _headerDefault.Append(wordParagraph._paragraph);
            //} else if (_type == HeaderFooterValues.Even) {
            //    _headerEven.Append(wordParagraph._paragraph);
            //}
            _header.Append(wordParagraph._paragraph);
            //this.Paragraphs.Add(wordParagraph);
            return wordParagraph;
        }

        public WordPageNumber AddPageNumber(WordPageNumberStyle wordPageNumberStyle) {
            var pageNumber = new WordPageNumber(_document, this, wordPageNumberStyle);
            return pageNumber;
        }
        public static void RemoveHeaders(WordprocessingDocument wordprocessingDocument) {
            var docPart = wordprocessingDocument.MainDocumentPart;
            DocumentFormat.OpenXml.Wordprocessing.Document document = docPart.Document;
            if (docPart.HeaderParts.Count() > 0) {
                // Remove the header
                docPart.DeleteParts(docPart.HeaderParts);

                // First, create a list of all descendants of type
                // HeaderReference. Then, navigate the list and call
                // Remove on each item to delete the reference.
                var headers = document.Descendants<HeaderReference>().ToList();
                foreach (var header in headers) {
                    header.Remove();
                }
            }
        }
    }
}
