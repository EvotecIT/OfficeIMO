using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class WordTabChar {
        private WordDocument _document;
        private readonly Paragraph _paragraph;
        private readonly Run _run;

        public WordTabChar(WordDocument document, Paragraph paragraph, Run run) {
            this._document = document;
            this._paragraph = paragraph;
            this._run = run;
        }

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

        internal static WordParagraph AddTab(WordDocument document, WordParagraph wordParagraph) {
            var newWordParagraph = new WordParagraph(document, wordParagraph._paragraph, true);
            newWordParagraph._run.Append(new TabChar());
            return newWordParagraph;
        }
    }
}
