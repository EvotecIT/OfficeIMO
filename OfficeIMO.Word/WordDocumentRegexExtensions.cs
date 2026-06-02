using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Shared;
using OfficeIMO.Word.Fluent;
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
        /// <returns>A <see cref="WordFind"/> instance containing all matches.</returns>
        public static WordFind Find(this WordDocument document, Regex regex) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }
            if (regex == null) {
                throw new ArgumentNullException(nameof(regex));
            }

            var result = new WordFind();

            result.FindRegex(document.Paragraphs, regex, result.Paragraphs);

            foreach (var table in document.Tables) {
                result.FindRegex(table.Paragraphs, regex, result.Tables);
            }

            if (document.Header?.Default != null) {
                result.FindRegex(document.Header.Default.Paragraphs, regex, result.HeaderDefault);
                foreach (var table in document.Header.Default.Tables) {
                    result.FindRegex(table.Paragraphs, regex, result.HeaderDefault);
                }
            }

            if (document.Header?.Even != null) {
                result.FindRegex(document.Header.Even.Paragraphs, regex, result.HeaderEven);
                foreach (var table in document.Header.Even.Tables) {
                    result.FindRegex(table.Paragraphs, regex, result.HeaderEven);
                }
            }

            if (document.Header?.First != null) {
                result.FindRegex(document.Header.First.Paragraphs, regex, result.HeaderFirst);
                foreach (var table in document.Header.First.Tables) {
                    result.FindRegex(table.Paragraphs, regex, result.HeaderFirst);
                }
            }

            if (document.Footer?.Default != null) {
                result.FindRegex(document.Footer.Default.Paragraphs, regex, result.FooterDefault);
                foreach (var table in document.Footer.Default.Tables) {
                    result.FindRegex(table.Paragraphs, regex, result.FooterDefault);
                }
            }

            if (document.Footer?.Even != null) {
                result.FindRegex(document.Footer.Even.Paragraphs, regex, result.FooterEven);
                foreach (var table in document.Footer.Even.Tables) {
                    result.FindRegex(table.Paragraphs, regex, result.FooterEven);
                }
            }

            if (document.Footer?.First != null) {
                result.FindRegex(document.Footer.First.Paragraphs, regex, result.FooterFirst);
                foreach (var table in document.Footer.First.Tables) {
                    result.FindRegex(table.Paragraphs, regex, result.FooterFirst);
                }
            }

            return result;
        }
    }
}
