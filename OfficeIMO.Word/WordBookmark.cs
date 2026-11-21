using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a bookmark within a Word document.
    /// </summary>
    public class WordBookmark : WordElement {
        private WordDocument _document;
        private WordParagraph? _authoringParagraph;

        /// <summary>Gets the paragraph containing the bookmark.</summary>
        public WordParagraph Paragraph => _authoringParagraph ?? new WordParagraph(_document, _paragraph);
        private Paragraph _paragraph;
        private BookmarkStart _bookmarkStart;

        private BookmarkEnd? _bookmarkEnd {
            get {
                foreach (var bookmarkEnd in _paragraph.Descendants<BookmarkEnd>()) {
                    if (bookmarkEnd.Id == _bookmarkStart.Id) {
                        return bookmarkEnd;
                    }
                }

                OpenXmlElement root = _bookmarkStart;
                while (root.Parent != null) {
                    root = root.Parent;
                }

                foreach (var bookmarkEnd in root.Descendants<BookmarkEnd>()) {
                    if (bookmarkEnd.Id == _bookmarkStart.Id) {
                        return bookmarkEnd;
                    }
                }

                return null;
            }
        }

        /// <summary>
        /// Gets or sets the bookmark name.
        /// </summary>
        public string? Name {
            get => _bookmarkStart.Name;
            set => _bookmarkStart.Name = value;
        }

        /// <summary>
        /// Gets or sets the bookmark identifier.
        /// </summary>
        public int Id {
            get => int.TryParse(_bookmarkStart.Id, out var id) ? id : 0;
            set => _bookmarkStart.Id = value.ToString();
        }

        /// <summary>
        /// Gets or sets the text the bookmark spans. Paragraph breaks are returned as <c>\n</c>.
        /// Setting the text replaces the bookmarked content and keeps the bookmark around the new text,
        /// so REF and cross-reference fields show it after their next update.
        /// </summary>
        /// <remarks>Setting the text follows <see cref="WordRange.Replace(string)"/>.</remarks>
        public string Text {
            get => GetRange().Text;
            set => GetRange().Replace(value ?? string.Empty);
        }

        /// <summary>Returns the range between the bookmark's start and end.</summary>
        /// <returns>The bookmarked range; empty when the bookmark marks a position rather than text.</returns>
        public WordRange GetRange() {
            if (_bookmarkStart.Parent == null) throw new InvalidOperationException("The bookmark was removed from the document.");
            (Paragraph startParagraph, int start) = LocateMarker(_bookmarkStart, isStart: true);
            BookmarkEnd? end = _bookmarkEnd;
            (Paragraph endParagraph, int endOffset) = end != null ? LocateMarker(end, isStart: false) : (startParagraph, start);
            bool endFirst = ReferenceEquals(startParagraph, endParagraph)
                ? endOffset < start
                : WordRange.GetParagraphsBetween(startParagraph, endParagraph) == null;
            if (endFirst) (endParagraph, endOffset) = (startParagraph, start);
            return new WordRange(_document, startParagraph, start, endParagraph, endOffset, _bookmarkStart);
        }

        private static (Paragraph Paragraph, int Offset) LocateMarker(OpenXmlElement marker, bool isStart) {
            Paragraph? paragraph = marker.Ancestors<Paragraph>().FirstOrDefault();
            if (paragraph != null) {
                var positions = new Dictionary<OpenXmlElement, int>();
                int index = 0;
                foreach (OpenXmlElement element in paragraph.Descendants()) positions[element] = index++;
                int markerPosition = positions[marker];
                int offset = 0;
                foreach (WordTextAtom atom in WordTextMap.Build(paragraph).Atoms) {
                    if (positions[atom.Element] < markerPosition) offset = atom.End;
                }
                return (paragraph, offset);
            }

            // A marker between paragraphs starts at the next paragraph or ends at the previous one.
            for (OpenXmlElement current = marker; current.Parent != null; current = current.Parent) {
                IEnumerable<OpenXmlElement> siblings = isStart ? current.ElementsAfter() : current.ElementsBefore().Reverse();
                foreach (OpenXmlElement sibling in siblings) {
                    Paragraph? found = sibling as Paragraph ?? (isStart
                        ? sibling.Descendants<Paragraph>().FirstOrDefault()
                        : sibling.Descendants<Paragraph>().LastOrDefault());
                    if (found != null) return (found, isStart ? 0 : WordTextMap.Build(found).Text.Length);
                }
                if (current.Parent is Body or Header or Footer or Footnote or Endnote or Comment) break;
            }
            throw new InvalidOperationException("The bookmark is not next to any paragraph.");
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordBookmark"/> class.
        /// </summary>
        /// <param name="document">Parent document.</param>
        /// <param name="paragraph">Paragraph containing the bookmark.</param>
        /// <param name="bookmarkStart">Underlying bookmark start element.</param>
        internal WordBookmark(WordDocument document, Paragraph paragraph, BookmarkStart bookmarkStart) {
            this._document = document;
            this._paragraph = paragraph;
            this._bookmarkStart = bookmarkStart;
        }

        /// <summary>
        /// Removes the bookmark from the document.
        /// </summary>
        public void Remove() {
            this._bookmarkEnd?.Remove();
            this._bookmarkStart.Remove();
        }

        /// <summary>
        /// Adds a bookmark to the specified paragraph.
        /// </summary>
        /// <param name="paragraph">Paragraph to contain the bookmark.</param>
        /// <param name="bookmarkName">Name of the bookmark.</param>
        /// <returns>The created bookmark.</returns>
        public static WordBookmark AddBookmark(WordParagraph paragraph, string bookmarkName) {
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
            return new WordBookmark(paragraph._document, paragraph._paragraph, bms) { _authoringParagraph = paragraph };
        }
    }
}
