using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Word {
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

        public string Name {
            get => _bookmarkStart.Name;
            set => _bookmarkStart.Name = value;
        }

        public int Id {
            get => int.Parse(_bookmarkStart.Id);
            set => _bookmarkStart.Id = value.ToString();
        }

        public WordBookmark(WordDocument document, Paragraph paragraph, BookmarkStart bookmarkStart) {
            this._document = document;
            this._paragraph = paragraph;
            this._bookmarkStart = bookmarkStart;
        }

        public void Remove() {
            this._bookmarkEnd.Remove();
            this._bookmarkStart.Remove();
        }

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
