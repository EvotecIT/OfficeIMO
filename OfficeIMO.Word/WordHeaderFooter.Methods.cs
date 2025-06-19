using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordHeaderFooter {
        /// <summary>
        /// Adds a paragraph with the specified text to the header or footer.
        /// </summary>
        public WordParagraph AddParagraph(string text) {
            var paragraph = AddParagraph();
            paragraph.Text = text;
            return paragraph;
        }

        /// <summary>
        /// Creates an empty paragraph in the header or footer. When
        /// <paramref name="newRun"/> is true an empty run is added.
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
        /// Adds a hyperlink to the header or footer pointing to an external URI.
        /// </summary>
        public WordParagraph AddHyperLink(string text, Uri uri, bool addStyle = false, string tooltip = "", bool history = true) {
            return this.AddParagraph().AddHyperLink(text, uri, addStyle, tooltip, history);
        }

        /// <summary>
        /// Adds a hyperlink that targets an internal bookmark within the document.
        /// </summary>
        public WordParagraph AddHyperLink(string text, string anchor, bool addStyle = false, string tooltip = "", bool history = true) {
            return this.AddParagraph().AddHyperLink(text, anchor, addStyle, tooltip, history);
        }

        /// <summary>
        /// Inserts a horizontal line into the header or footer.
        /// </summary>
        public WordParagraph AddHorizontalLine(BorderValues? lineType = null, SixLabors.ImageSharp.Color? color = null, uint size = 12, uint space = 1) {
            lineType ??= BorderValues.Single;
            return this.AddParagraph().AddHorizontalLine(lineType.Value, color, size, space);
        }

        /// <summary>
        /// Adds a bookmark at the current location in the header or footer.
        /// </summary>
        public WordParagraph AddBookmark(string bookmarkName) {
            return this.AddParagraph().AddBookmark(bookmarkName);
        }

        /// <summary>
        /// Inserts a field into the header or footer using the specified type and format.
        /// </summary>
        public WordParagraph AddField(WordFieldType wordFieldType, WordFieldFormat? wordFieldFormat = null, bool advanced = false) {
            return this.AddParagraph().AddField(wordFieldType, wordFieldFormat, advanced);
        }

        /// <summary>
        /// Adds a table to the header or footer with the given dimensions and style.
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
        /// Creates a list in the header or footer and applies the specified style.
        /// </summary>
        public WordList AddList(WordListStyle style) {
            WordList wordList = new WordList(this._document, this);
            wordList.AddList(style);
            return wordList;
        }
    }
}
