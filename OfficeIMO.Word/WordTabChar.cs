using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a tab character within a paragraph.
    /// </summary>
    public class WordTabChar : WordElement {
        private WordDocument _document;
        private readonly Paragraph _paragraph;
        private readonly Run _run;

        /// <summary>
        /// Initializes a new instance of the <see cref="WordTabChar"/> class.
        /// </summary>
        /// <param name="document">Parent document.</param>
        /// <param name="paragraph">Paragraph containing the tab.</param>
        /// <param name="run">Run that holds the tab character.</param>
        public WordTabChar(WordDocument document, Paragraph paragraph, Run run) {
            this._document = document;
            this._paragraph = paragraph;
            this._run = run;
        }

        /// <summary>
        /// Removes the tab character from the document. Optionally removes the paragraph.
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
        /// Appends a tab character to the specified paragraph.
        /// </summary>
        /// <param name="document">Parent document.</param>
        /// <param name="wordParagraph">Paragraph to add the tab to.</param>
        /// <returns>A paragraph containing the inserted tab.</returns>
        internal static WordParagraph AddTab(WordDocument document, WordParagraph wordParagraph) {
            var newWordParagraph = new WordParagraph(document, wordParagraph._paragraph, true);
            newWordParagraph._run.Append(new TabChar());
            return newWordParagraph;
        }
    }
}
