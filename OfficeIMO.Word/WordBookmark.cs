using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a bookmark within a Word document.
    /// </summary>
    public class WordBookmark : WordElement {
        private WordDocument _document;
        private Paragraph _paragraph;
        private BookmarkStart _bookmarkStart;

        private BookmarkEnd _bookmarkEnd {
            get {
                var listElements = _document._wordprocessingDocument.MainDocumentPart.Document.Body.ChildElements.OfType<Paragraph>();
                foreach (Paragraph paragraph in listElements) {
                    var listBookmarkEnds = paragraph.ChildElements.OfType<BookmarkEnd>();
                    foreach (var bookmarkEnd in listBookmarkEnds) {
                        if (bookmarkEnd.Id == _bookmarkStart.Id) {
                            return bookmarkEnd;
                        }
                    }
                }

                return null;
            }
        }

        /// <summary>
        /// Gets or sets the bookmark name.
        /// </summary>
        public string Name {
            get => _bookmarkStart.Name;
            set => _bookmarkStart.Name = value;
        }

        /// <summary>
        /// Gets or sets the bookmark identifier.
        /// </summary>
        public int Id {
            get => int.Parse(_bookmarkStart.Id);
            set => _bookmarkStart.Id = value.ToString();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordBookmark"/> class.
        /// </summary>
        /// <param name="document">Parent document.</param>
        /// <param name="paragraph">Paragraph containing the bookmark.</param>
        /// <param name="bookmarkStart">Underlying bookmark start element.</param>
        public WordBookmark(WordDocument document, Paragraph paragraph, BookmarkStart bookmarkStart) {
            this._document = document;
            this._paragraph = paragraph;
            this._bookmarkStart = bookmarkStart;
        }

        /// <summary>
        /// Removes the bookmark from the document.
        /// </summary>
        public void Remove() {
            this._bookmarkEnd.Remove();
            this._bookmarkStart.Remove();
        }

        /// <summary>
        /// Adds a bookmark to the specified paragraph.
        /// </summary>
        /// <param name="paragraph">Paragraph to contain the bookmark.</param>
        /// <param name="bookmarkName">Name of the bookmark.</param>
        /// <returns>The paragraph with the inserted bookmark.</returns>
        public static WordParagraph AddBookmark(WordParagraph paragraph, string bookmarkName) {
            BookmarkStart bms = new BookmarkStart() { Name = bookmarkName, Id = paragraph._document.BookmarkId.ToString() };
            BookmarkEnd bme = new BookmarkEnd() { Id = paragraph._document.BookmarkId.ToString() };

            //paragraph.VerifyRun();
            if (paragraph._run == null) {
                paragraph._paragraph.Append(bms);
                paragraph._paragraph.Append(bme);
            } else {
                var bm = paragraph._run.InsertAfterSelf(bms);
                bm.InsertAfterSelf(bme);
            }


            paragraph._bookmarkStart = bms;
            return paragraph;
        }
    }
}
