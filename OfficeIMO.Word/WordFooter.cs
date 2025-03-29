using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class WordFooters {
        public WordFooter Default {
            get;
            set;
        }

        public WordFooter Even {
            get;
            set;
        }

        public WordFooter First {
            get;
            set;
        }
    }
    public partial class WordFooter : WordHeaderFooter {
        private readonly WordSection _section;

        internal WordFooter(WordDocument document, FooterReference footerReference, WordSection section) {
            _document = document;
            _id = footerReference.Id;
            _type = WordSection.GetType(footerReference.Type);
            _section = section;

            var listHeaders = document._wordprocessingDocument.MainDocumentPart.FooterParts.ToList();
            foreach (FooterPart footerPart in listHeaders) {
                var id = document._wordprocessingDocument.MainDocumentPart.GetIdOfPart(footerPart);
                if (id == _id) {
                    _footerPart = footerPart;
                    _footer = footerPart.Footer;
                }
            }

            if (_type == HeaderFooterValues.Default) {
                document._currentSection.Footer.Default = this;
            } else if (_type == HeaderFooterValues.Even) {
                document._currentSection.Footer.Even = this;
            } else if (_type == HeaderFooterValues.First) {
                document._currentSection.Footer.First = this;
            } else {
                throw new InvalidOperationException("Shouldn't happen?");
            }
        }

        internal WordFooter(WordDocument document, HeaderFooterValues type, Footer footerPartFooter, WordSection section) {
            _document = document;
            _footer = footerPartFooter;
            _type = type;
            _section = section;
        }

        public WordPageNumber AddPageNumber(WordPageNumberStyle wordPageNumberStyle) {
            var pageNumber = new WordPageNumber(_document, this, wordPageNumberStyle);
            return pageNumber;
        }

        public static void RemoveFooters(WordprocessingDocument wordprocessingDocument) {
            var docPart = wordprocessingDocument.MainDocumentPart;
            DocumentFormat.OpenXml.Wordprocessing.Document document = docPart.Document;
            if (docPart.FooterParts.Any()) {
                // Remove the header
                docPart.DeleteParts(docPart.FooterParts);

                // First, create a list of all descendants of type
                // HeaderReference. Then, navigate the list and call
                // Remove on each item to delete the reference.
                var footers = document.Descendants<FooterReference>().ToList();
                foreach (var footer in footers) {
                    footer.Remove();
                }
            }
        }
        public static void RemoveFooters(WordDocument document) {
            RemoveFooters(document._wordprocessingDocument);
        }
    }
}
