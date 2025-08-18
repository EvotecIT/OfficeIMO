using System;
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

        internal void ForEachParagraph(Action<WordParagraph> action) {
            foreach (var paragraph in EnumerateAllParagraphs()) {
                action(paragraph);
            }
        }

        internal IEnumerable<WordParagraph> FindParagraphs(string text, StringComparison stringComparison = StringComparison.OrdinalIgnoreCase) {
            foreach (var paragraph in EnumerateAllParagraphs()) {
                if (paragraph.Text?.IndexOf(text, stringComparison) >= 0) {
                    yield return paragraph;
                }
            }
        }

        internal IEnumerable<WordParagraph> SelectParagraphs(Func<WordParagraph, bool> predicate) {
            foreach (var paragraph in EnumerateAllParagraphs()) {
                if (predicate(paragraph)) {
                    yield return paragraph;
                }
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
