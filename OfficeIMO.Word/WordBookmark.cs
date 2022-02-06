using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Word {
    public class WordBookmark {
        private WordDocument _document;
        private Paragraph _paragraph;
        private readonly BookmarkStart _bookmarkStart;
        private BookmarkEnd _bookmarkEnd;

        public string Name {
            get => _bookmarkStart.Name;
            set => _bookmarkStart.Name = value;
        }

        public string Id {
            get => _bookmarkStart.Id;
            set => _bookmarkStart.Id = value;
        }

        public WordBookmark(WordDocument document, Paragraph paragraph, BookmarkStart bookmarkStart) {
            this._document = document;
            this._paragraph = paragraph;
            this._bookmarkStart = bookmarkStart;
        }
    }
}
