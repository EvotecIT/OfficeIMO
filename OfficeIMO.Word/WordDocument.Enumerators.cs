using System.Text.RegularExpressions;

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

            // Iterate headers/footers per section to avoid multi-section header warnings
            foreach (var sec in this.Sections) {
                var headers = sec.Header;
                foreach (var p in EnumerateHeaderFooterParagraphs(headers.Default)) yield return p;
                foreach (var p in EnumerateHeaderFooterParagraphs(headers.Even)) yield return p;
                foreach (var p in EnumerateHeaderFooterParagraphs(headers.First)) yield return p;

                var footers = sec.Footer;
                foreach (var p in EnumerateHeaderFooterParagraphs(footers.Default)) yield return p;
                foreach (var p in EnumerateHeaderFooterParagraphs(footers.Even)) yield return p;
                foreach (var p in EnumerateHeaderFooterParagraphs(footers.First)) yield return p;
            }
        }

        internal void ForEachParagraph(Action<WordParagraph> action) {
            foreach (var paragraph in EnumerateAllParagraphs()) {
                action(paragraph);
            }
        }

        internal void ForEachRun(Action<WordParagraph> action) {
            foreach (var paragraph in EnumerateAllParagraphs()) {
                foreach (var run in paragraph.GetRuns()) {
                    action(run);
                }
            }
        }

        internal IEnumerable<WordParagraph> FindParagraphs(string text, StringComparison stringComparison = StringComparison.OrdinalIgnoreCase) {
            foreach (var paragraph in EnumerateAllParagraphs()) {
                if (paragraph.Text?.IndexOf(text, stringComparison) >= 0) {
                    yield return paragraph;
                }
            }
        }

        internal IEnumerable<WordParagraph> FindRunsRegex(string pattern) {
            var regex = new Regex(pattern);
            foreach (var paragraph in EnumerateAllParagraphs()) {
                foreach (var run in paragraph.GetRuns()) {
                    if (run.Text != null && regex.IsMatch(run.Text)) {
                        yield return run;
                    }
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

        private static IEnumerable<WordParagraph> EnumerateHeaderFooterParagraphs(WordHeaderFooter? headerFooter) {
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
