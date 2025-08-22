using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides access to the footer instances associated with a
    /// <see cref="WordDocument"/> section.
    /// </summary>
    public class WordFooters {
        /// <summary>
        /// Gets or sets the default footer for the section.
        /// </summary>
        public WordFooter? Default {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the footer used for even pages.
        /// </summary>
        public WordFooter? Even {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the footer used for the first page.
        /// </summary>
        public WordFooter? First {
            get;
            set;
        }
    }
    /// <summary>
    /// Represents a footer in a Word document and allows manipulation
    /// of its contents.
    /// </summary>
    public partial class WordFooter : WordHeaderFooter {
        private readonly WordSection _section;

        internal WordFooter(WordDocument document, FooterReference footerReference, WordSection section) {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _id = footerReference.Id?.Value ?? throw new InvalidOperationException("FooterReference requires an Id.");
            var typeEnum = footerReference.Type?.Value;
            var typeString = typeEnum?.ToString().ToLower() ?? "first";
            _type = WordSection.GetType(typeString);
            _section = section;

            var mainPart = document._wordprocessingDocument.MainDocumentPart
                ?? throw new InvalidOperationException("Document does not contain a MainDocumentPart.");
            var listHeaders = mainPart.FooterParts.ToList();
            foreach (FooterPart footerPart in listHeaders) {
                var id = mainPart.GetIdOfPart(footerPart);
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
        /// Adds a page number to this footer using the specified style.
        /// </summary>
        /// <param name="wordPageNumberStyle">Style of the page number to insert.</param>
        /// <returns>The created <see cref="WordPageNumber"/> instance.</returns>
        public WordPageNumber AddPageNumber(WordPageNumberStyle wordPageNumberStyle) {
            var pageNumber = new WordPageNumber(_document, this, wordPageNumberStyle);
            return pageNumber;
        }

        /// <summary>
        /// Removes footers from the provided <see cref="WordprocessingDocument"/>.
        /// When no <paramref name="types"/> are specified all footers are removed.
        /// </summary>
        /// <param name="wordprocessingDocument">Document to operate on.</param>
        /// <param name="types">Footer types to remove.</param>
        public static void RemoveFooters(WordprocessingDocument wordprocessingDocument, params HeaderFooterValues[] types) {
            var docPart = wordprocessingDocument.MainDocumentPart
                ?? throw new InvalidOperationException("WordprocessingDocument does not contain a MainDocumentPart.");
            var document = docPart.Document
                ?? throw new InvalidOperationException("MainDocumentPart does not contain a Document.");

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
                .Where(f => f.Type != null && types.Contains(f.Type))
                .ToList();
            foreach (var footer in footersToRemove) {
                var footerId = footer.Id?.Value;
                if (footerId != null) {
                    var part = docPart.GetPartById(footerId) as FooterPart;
                    if (part != null) {
                        partsToDelete.Add(part);
                    }
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