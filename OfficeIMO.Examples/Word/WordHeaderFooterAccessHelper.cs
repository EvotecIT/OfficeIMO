using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static class WordHeaderFooterAccessHelper {
        internal static WordHeader GetSectionHeaderOrThrow(WordSection section, HeaderFooterValues type) {
            if (section == null) {
                throw new ArgumentNullException(nameof(section));
            }

            var header = section.GetHeader(type);
            if (header != null) {
                return header;
            }

            section.AddHeadersAndFooters();
            header = section.GetHeader(type);
            if (header == null) {
                throw new InvalidOperationException($"Section header '{type}' is not available.");
            }

            return header;
        }

        internal static WordFooter GetSectionFooterOrThrow(WordSection section, HeaderFooterValues type) {
            if (section == null) {
                throw new ArgumentNullException(nameof(section));
            }

            var footer = section.GetFooter(type);
            if (footer != null) {
                return footer;
            }

            section.AddHeadersAndFooters();
            footer = section.GetFooter(type);
            if (footer == null) {
                throw new InvalidOperationException($"Section footer '{type}' is not available.");
            }

            return footer;
        }

        internal static WordHeader GetDocumentHeaderOrThrow(WordDocument document, HeaderFooterValues type) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            var headers = document.Header;
            if (headers == null) {
                document.AddHeadersAndFooters();
                headers = document.Header;
            }

            if (headers == null) {
                throw new InvalidOperationException("Document headers are not available after AddHeadersAndFooters().");
            }

            WordHeader? header;
            if (type == HeaderFooterValues.First) {
                header = headers.First;
            } else if (type == HeaderFooterValues.Even) {
                header = headers.Even;
            } else {
                header = headers.Default;
            }

            if (header != null) {
                return header;
            }

            document.AddHeadersAndFooters();
            headers = document.Header ?? throw new InvalidOperationException("Document headers are not available after AddHeadersAndFooters().");
            if (type == HeaderFooterValues.First) {
                header = headers.First;
            } else if (type == HeaderFooterValues.Even) {
                header = headers.Even;
            } else {
                header = headers.Default;
            }

            if (header == null) {
                throw new InvalidOperationException($"Document header '{type}' is not available.");
            }

            return header;
        }

        internal static WordFooter GetDocumentFooterOrThrow(WordDocument document, HeaderFooterValues type) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            var footers = document.Footer;
            if (footers == null) {
                document.AddHeadersAndFooters();
                footers = document.Footer;
            }

            if (footers == null) {
                throw new InvalidOperationException("Document footers are not available after AddHeadersAndFooters().");
            }

            WordFooter? footer;
            if (type == HeaderFooterValues.First) {
                footer = footers.First;
            } else if (type == HeaderFooterValues.Even) {
                footer = footers.Even;
            } else {
                footer = footers.Default;
            }

            if (footer != null) {
                return footer;
            }

            document.AddHeadersAndFooters();
            footers = document.Footer ?? throw new InvalidOperationException("Document footers are not available after AddHeadersAndFooters().");
            if (type == HeaderFooterValues.First) {
                footer = footers.First;
            } else if (type == HeaderFooterValues.Even) {
                footer = footers.Even;
            } else {
                footer = footers.Default;
            }

            if (footer == null) {
                throw new InvalidOperationException($"Document footer '{type}' is not available.");
            }

            return footer;
        }
    }
}
