using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordHeaderFooter {
        /// <summary>
        /// Executes the AddParagraph operation.
        /// </summary>
        public WordParagraph AddParagraph(string text) {
            var paragraph = AddParagraph();
            paragraph.Text = text;
            return paragraph;
        }

        /// <summary>
        /// Executes the AddParagraph operation.
        /// </summary>
        public WordParagraph AddParagraph(bool newRun = false) {
            var wordParagraph = new WordParagraph(_document, newParagraph: true, newRun: newRun);
            if (_footer != null) {
                _footer.Append(wordParagraph._paragraph);
            } else if (_header != null) {
                _header.Append(wordParagraph._paragraph);
            }
            return wordParagraph;
        }

        /// <summary>
        /// Executes the AddHyperLink operation.
        /// </summary>
        public WordParagraph AddHyperLink(string text, Uri uri, bool addStyle = false, string tooltip = "", bool history = true) {
            return this.AddParagraph().AddHyperLink(text, uri, addStyle, tooltip, history);
        }

        /// <summary>
        /// Executes the AddHyperLink operation.
        /// </summary>
        public WordParagraph AddHyperLink(string text, string anchor, bool addStyle = false, string tooltip = "", bool history = true) {
            return this.AddParagraph().AddHyperLink(text, anchor, addStyle, tooltip, history);
        }

        /// <summary>
        /// Executes the AddHorizontalLine operation.
        /// </summary>
        public WordParagraph AddHorizontalLine(BorderValues? lineType = null, SixLabors.ImageSharp.Color? color = null, uint size = 12, uint space = 1) {
            lineType ??= BorderValues.Single;
            return this.AddParagraph().AddHorizontalLine(lineType.Value, color, size, space);
        }

        /// <summary>
        /// Executes the AddBookmark operation.
        /// </summary>
        public WordParagraph AddBookmark(string bookmarkName) {
            return this.AddParagraph().AddBookmark(bookmarkName);
        }

        /// <summary>
        /// Executes the AddField operation.
        /// </summary>
        public WordParagraph AddField(WordFieldType wordFieldType, WordFieldFormat? wordFieldFormat = null, bool advanced = false) {
            return this.AddParagraph().AddField(wordFieldType, wordFieldFormat, advanced);
        }

        /// <summary>
        /// Executes the AddTable operation.
        /// </summary>
        public WordTable AddTable(int rows, int columns, WordTableStyle tableStyle = WordTableStyle.TableGrid) {
            if (_footer != null) {
                return new WordTable(_document, _footer, rows, columns, tableStyle);
            } else if (_header != null) {
                return new WordTable(_document, _header, rows, columns, tableStyle);
            } else {
                throw new InvalidOperationException("No footer or header defined. That is weird.");
            }
        }

        /// <summary>
        /// Executes the AddList operation.
        /// </summary>
        public WordList AddList(WordListStyle style) {
            WordList wordList = new WordList(this._document, this);
            wordList.AddList(style);
            return wordList;
        }
    }
}
