using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Core.Internal;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;


namespace OfficeIMO.Word {
    /// <summary>
    /// Extension methods for <see cref="WordDocument"/> providing regular expression search.
    /// </summary>
    public static class WordDocumentRegexExtensions {
        /// <summary>
        /// Searches the document for text matching the specified regular expression.
        /// </summary>
        /// <param name="document">Document to search.</param>
        /// <param name="regex">Regular expression used for searching.</param>
        /// <returns>A <see cref="WordSearchResult"/> containing all matches.</returns>
        public static WordSearchResult Find(this WordDocument document, Regex regex) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }
            if (regex == null) {
                throw new ArgumentNullException(nameof(regex));
            }

            var result = new WordSearchResult();

            result.FindRegex(document.Paragraphs, regex, result.MutableBodyParagraphs);

            foreach (var table in document.Tables) {
                result.FindRegex(table.Paragraphs, regex, result.MutableTableParagraphs);
            }

            var headers = document.Sections[0].Headers;
            var footers = document.Sections[0].Footers;
            if (headers.Default is { } defaultHeader) {
                result.FindRegex(defaultHeader.Paragraphs, regex, result.MutableDefaultHeaderParagraphs);
                foreach (var table in defaultHeader.Tables) {
                    result.FindRegex(table.Paragraphs, regex, result.MutableDefaultHeaderParagraphs);
                }
            }

            if (headers.Even is { } evenHeader) {
                result.FindRegex(evenHeader.Paragraphs, regex, result.MutableEvenPageHeaderParagraphs);
                foreach (var table in evenHeader.Tables) {
                    result.FindRegex(table.Paragraphs, regex, result.MutableEvenPageHeaderParagraphs);
                }
            }

            if (headers.First is { } firstHeader) {
                result.FindRegex(firstHeader.Paragraphs, regex, result.MutableFirstPageHeaderParagraphs);
                foreach (var table in firstHeader.Tables) {
                    result.FindRegex(table.Paragraphs, regex, result.MutableFirstPageHeaderParagraphs);
                }
            }

            if (footers.Default is { } defaultFooter) {
                result.FindRegex(defaultFooter.Paragraphs, regex, result.MutableDefaultFooterParagraphs);
                foreach (var table in defaultFooter.Tables) {
                    result.FindRegex(table.Paragraphs, regex, result.MutableDefaultFooterParagraphs);
                }
            }

            if (footers.Even is { } evenFooter) {
                result.FindRegex(evenFooter.Paragraphs, regex, result.MutableEvenPageFooterParagraphs);
                foreach (var table in evenFooter.Tables) {
                    result.FindRegex(table.Paragraphs, regex, result.MutableEvenPageFooterParagraphs);
                }
            }

            if (footers.First is { } firstFooter) {
                result.FindRegex(firstFooter.Paragraphs, regex, result.MutableFirstPageFooterParagraphs);
                foreach (var table in firstFooter.Tables) {
                    result.FindRegex(table.Paragraphs, regex, result.MutableFirstPageFooterParagraphs);
                }
            }

            return result;
        }
    }
}
