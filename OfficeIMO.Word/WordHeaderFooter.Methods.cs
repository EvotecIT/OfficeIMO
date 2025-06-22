using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordHeaderFooter {
        /// <summary>
        /// Adds a new paragraph with the specified text to the header or footer.
        /// </summary>
        /// <param name="text">Paragraph text.</param>
        /// <returns>The created <see cref="WordParagraph"/> instance.</returns>
        public WordParagraph AddParagraph(string text) {
            var paragraph = AddParagraph();
            paragraph.Text = text;
            return paragraph;
        }

        /// <summary>
        /// Creates an empty paragraph and appends it to the header or footer.
        /// </summary>
        /// <param name="newRun">Specifies whether a new run should be started.</param>
        /// <returns>The created <see cref="WordParagraph"/> instance.</returns>
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
        /// Adds a hyperlink pointing to an external URI.
        /// </summary>
        /// <param name="text">Display text for the hyperlink.</param>
        /// <param name="uri">Destination URI.</param>
        /// <param name="addStyle">Whether to apply hyperlink style.</param>
        /// <param name="tooltip">Tooltip text for the link.</param>
        /// <param name="history">Whether to mark the link as visited.</param>
        /// <returns>The created <see cref="WordParagraph"/> instance.</returns>
        public WordParagraph AddHyperLink(string text, Uri uri, bool addStyle = false, string tooltip = "", bool history = true) {
            return this.AddParagraph().AddHyperLink(text, uri, addStyle, tooltip, history);
        }

        /// <summary>
        /// Adds an internal hyperlink to a bookmark within the document.
        /// </summary>
        /// <param name="text">Display text for the hyperlink.</param>
        /// <param name="anchor">Bookmark to link to.</param>
        /// <param name="addStyle">Whether to apply hyperlink style.</param>
        /// <param name="tooltip">Tooltip text for the link.</param>
        /// <param name="history">Whether to mark the link as visited.</param>
        /// <returns>The created <see cref="WordParagraph"/> instance.</returns>
        public WordParagraph AddHyperLink(string text, string anchor, bool addStyle = false, string tooltip = "", bool history = true) {
            return this.AddParagraph().AddHyperLink(text, anchor, addStyle, tooltip, history);
        }

        /// <summary>
        /// Adds a horizontal line to the header or footer.
        /// </summary>
        /// <param name="lineType">Border style of the line.</param>
        /// <param name="color">Color of the line.</param>
        /// <param name="size">Thickness of the line.</param>
        /// <param name="space">Spacing above and below the line.</param>
        /// <returns>The created <see cref="WordParagraph"/> instance.</returns>
        public WordParagraph AddHorizontalLine(BorderValues? lineType = null, SixLabors.ImageSharp.Color? color = null, uint size = 12, uint space = 1) {
            lineType ??= BorderValues.Single;
            return this.AddParagraph().AddHorizontalLine(lineType.Value, color, size, space);
        }

        /// <summary>
        /// Adds a bookmark at the current location.
        /// </summary>
        /// <param name="bookmarkName">Name of the bookmark.</param>
        /// <returns>The created <see cref="WordParagraph"/> instance.</returns>
        public WordParagraph AddBookmark(string bookmarkName) {
            return this.AddParagraph().AddBookmark(bookmarkName);
        }

        /// <summary>
        /// Adds a field to the header or footer.
        /// </summary>
        /// <param name="wordFieldType">Type of field to insert.</param>
        /// <param name="wordFieldFormat">Optional field format.</param>
        /// <param name="advanced">Whether to use advanced formatting.</param>
        /// <returns>The created <see cref="WordParagraph"/> instance.</returns>
        public WordParagraph AddField(WordFieldType wordFieldType, WordFieldFormat? wordFieldFormat = null, bool advanced = false) {
            return this.AddParagraph().AddField(wordFieldType, wordFieldFormat, advanced);
        }

        /// <summary>
        /// Inserts a page number field.
        /// </summary>
        /// <param name="includeTotalPages">Include total pages in the display.</param>
        /// <param name="format">Optional number format.</param>
        /// <param name="separator">Separator used when total pages are included.</param>
        /// <returns>The created <see cref="WordParagraph"/> instance.</returns>
        public WordParagraph AddPageNumber(bool includeTotalPages = false, WordFieldFormat? format = null, string separator = " of ") {
            return this.AddParagraph().AddPageNumber(includeTotalPages, format, separator);
        }

        /// <summary>
        /// Creates a table and appends it to the header or footer.
        /// </summary>
        /// <param name="rows">Number of rows.</param>
        /// <param name="columns">Number of columns.</param>
        /// <param name="tableStyle">Table style to apply.</param>
        /// <returns>The created <see cref="WordTable"/>.</returns>
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
        /// Creates a list with the specified style.
        /// </summary>
        /// <param name="style">List style to apply.</param>
        /// <returns>The created <see cref="WordList"/>.</returns>
        public WordList AddList(WordListStyle style) {
            WordList wordList = new WordList(this._document, this);
            wordList.AddList(style);
            return wordList;
        }
    }
}
