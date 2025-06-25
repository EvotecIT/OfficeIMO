using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides enumerators for traversing document content.
    /// </summary>
    public partial class WordDocument {
        internal IEnumerable<WordParagraph> EnumerateAllParagraphs() {
            foreach (var paragraph in this.Paragraphs) {
                yield return paragraph;
            }

            foreach (var table in this.Tables) {
                foreach (var paragraph in table.Paragraphs) {
                    yield return paragraph;
                }
            }

            foreach (var paragraph in EnumerateHeaderFooterParagraphs(this.Header.Default)) {
                yield return paragraph;
            }
            foreach (var paragraph in EnumerateHeaderFooterParagraphs(this.Header.Even)) {
                yield return paragraph;
            }
            foreach (var paragraph in EnumerateHeaderFooterParagraphs(this.Header.First)) {
                yield return paragraph;
            }
            foreach (var paragraph in EnumerateHeaderFooterParagraphs(this.Footer.Default)) {
                yield return paragraph;
            }
            foreach (var paragraph in EnumerateHeaderFooterParagraphs(this.Footer.Even)) {
                yield return paragraph;
            }
            foreach (var paragraph in EnumerateHeaderFooterParagraphs(this.Footer.First)) {
                yield return paragraph;
            }
        }

        private static IEnumerable<WordParagraph> EnumerateHeaderFooterParagraphs(WordHeaderFooter headerFooter) {
            if (headerFooter == null) yield break;

            foreach (var paragraph in headerFooter.Paragraphs) {
                yield return paragraph;
            }

            foreach (var table in headerFooter.Tables) {
                foreach (var paragraph in table.Paragraphs) {
                    yield return paragraph;
                }
            }
        }
    }
}
