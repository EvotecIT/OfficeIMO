using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides operations for a single tab character within a paragraph.
    /// </summary>
    public class WordTabChar : WordElement {
        private WordDocument _document;
        private readonly Paragraph _paragraph;
        private readonly Run _run;

        /// <summary>
        /// Initializes a new instance of the <see cref="WordTabChar"/> class bound to the
        /// specified document and paragraph.
        /// </summary>
        /// <param name="document">Parent document.</param>
        /// <param name="paragraph">Paragraph containing the tab.</param>
        /// <param name="run">Run that holds the tab character.</param>
        public WordTabChar(WordDocument document, Paragraph paragraph, Run run) {
            this._document = document ?? throw new ArgumentNullException(nameof(document));
            this._paragraph = paragraph ?? throw new ArgumentNullException(nameof(paragraph));
            this._run = run ?? throw new ArgumentNullException(nameof(run));
        }

        /// <summary>
        /// Removes this tab character from the document and optionally deletes the parent paragraph.
        /// </summary>
        /// <param name="includingParagraph">
        /// If set to <c>true</c> the paragraph containing the tab is removed.
        /// </param>
        public void Remove(bool includingParagraph = false) {
            if (includingParagraph) {
                this._paragraph.Remove();
            } else {
                if (_run.ChildElements.Count == 1) {
                    this._run.Remove();
                } else {
                    this._run.ChildElements.OfType<TabChar>().FirstOrDefault()?.Remove();
                }
            }
        }

        /// <summary>
        /// Inserts a new tab character into a paragraph and returns the containing <see cref="WordParagraph"/>.
        /// </summary>
        /// <param name="document">Parent document.</param>
        /// <param name="wordParagraph">Paragraph to add the tab to.</param>
        /// <returns>A paragraph containing the inserted tab.</returns>
        internal static WordParagraph AddTab(WordDocument document, WordParagraph wordParagraph) {
            var newWordParagraph = new WordParagraph(document, wordParagraph._paragraph, true);
            var run = newWordParagraph.VerifyRun();
            run.Append(new TabChar());
            return newWordParagraph;
        }
    }
}
