using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordHeaderFooter {
        public WordParagraph AddParagraph(string text) {
            var paragraph = AddParagraph();
            paragraph.Text = text;
            return paragraph;
        }

        public WordParagraph AddParagraph(bool newRun = false) {
            var wordParagraph = new WordParagraph(_document, newParagraph: true, newRun: newRun);
            if (_footer != null) {
                _footer.Append(wordParagraph._paragraph);
            } else if (_header != null) {
                _header.Append(wordParagraph._paragraph);
            }
            return wordParagraph;
        }

        public WordParagraph AddHyperLink(string text, Uri uri, bool addStyle = false, string tooltip = "", bool history = true) {
            return this.AddParagraph().AddHyperLink(text, uri, addStyle, tooltip, history);
        }

        public WordParagraph AddHyperLink(string text, string anchor, bool addStyle = false, string tooltip = "", bool history = true) {
            return this.AddParagraph().AddHyperLink(text, anchor, addStyle, tooltip, history);
        }

        public WordParagraph AddHorizontalLine(BorderValues lineType = BorderValues.Single, SixLabors.ImageSharp.Color? color = null, uint size = 12, uint space = 1) {
            return this.AddParagraph().AddHorizontalLine(lineType, color, size, space);
        }

        public WordParagraph AddBookmark(string bookmarkName) {
            return this.AddParagraph().AddBookmark(bookmarkName);
        }

        public WordParagraph AddField(WordFieldType wordFieldType, WordFieldFormat? wordFieldFormat = null, bool advanced = false) {
            return this.AddParagraph().AddField(wordFieldType, wordFieldFormat, advanced);
        }

        public WordTable AddTable(int rows, int columns, WordTableStyle tableStyle = WordTableStyle.TableGrid) {
            if (_footer != null) {
                return new WordTable(_document, _footer, rows, columns, tableStyle);
            } else if (_header != null) {
                return new WordTable(_document, _header, rows, columns, tableStyle);
            } else {
                throw new InvalidOperationException("No footer or header defined. That is weird.");
            }
        }

        public WordList AddList(WordListStyle style, bool continueNumbering = false) {
            WordList wordList = new WordList(this._document, this);
            wordList.AddList(style, continueNumbering);
            return wordList;
        }
    }
}
