using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class WordFooters {
        /// <summary>
        /// Gets or sets the Default.
        /// </summary>
        public WordFooter Default {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the Even.
        /// </summary>
        public WordFooter Even {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the First.
        /// </summary>
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

        /// <summary>
        /// Inserts a page number field into this footer using the specified style.
        /// </summary>
        public WordPageNumber AddPageNumber(WordPageNumberStyle wordPageNumberStyle) {
            var pageNumber = new WordPageNumber(_document, this, wordPageNumberStyle);
            return pageNumber;
        }

        /// <summary>
        /// Removes all footer parts and references from the provided document.
        /// Removes all footers from the given <see cref="WordDocument"/> instance.
        /// </summary>
        /// <param name="wordprocessingDocument">Document to operate on.</param>
        /// <param name="types">Footer types to remove.</param>
        public static void RemoveFooters(WordprocessingDocument wordprocessingDocument, params HeaderFooterValues[] types) {
            var docPart = wordprocessingDocument.MainDocumentPart;
            DocumentFormat.OpenXml.Wordprocessing.Document document = docPart.Document;

            if (types == null || types.Length == 0) {
                if (docPart.FooterParts.Any()) {
                    docPart.DeleteParts(docPart.FooterParts);
                    var footers = document.Descendants<FooterReference>().ToList();
                    foreach (var footer in footers) {
                        footer.Remove();
                    }
                }
                return;
            }

            var partsToDelete = new HashSet<FooterPart>();
            var footersToRemove = document.Descendants<FooterReference>()
                .Where(f => types.Contains(f.Type)).ToList();
            foreach (var footer in footersToRemove) {
                var part = docPart.GetPartById(footer.Id) as FooterPart;
                if (part != null) {
                    partsToDelete.Add(part);
                }
                footer.Remove();
            }
            foreach (var part in partsToDelete) {
                docPart.DeletePart(part);
            }
        }
        /// <summary>
        /// Removes footers from the specified <see cref="WordDocument"/>.
        /// When no <paramref name="types"/> are provided all footers are removed.
        /// </summary>
        /// <param name="document">Document to operate on.</param>
        /// <param name="types">Footer types to remove.</param>
        public static void RemoveFooters(WordDocument document, params HeaderFooterValues[] types) {
            RemoveFooters(document._wordprocessingDocument, types);
        }
    }
}
