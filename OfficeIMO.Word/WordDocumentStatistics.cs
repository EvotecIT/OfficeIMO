namespace OfficeIMO.Word {
    /// <summary>
    /// Provides basic statistics for a <see cref="WordDocument"/> such as page, paragraph,
    /// word and image counts. Additional statistics like table or chart counts
    /// are also available.
    /// </summary>
    public class WordDocumentStatistics {
        private readonly WordDocument _document;

        /// <summary>
        /// Initializes a new instance of the <see cref="WordDocumentStatistics"/> class.
        /// </summary>
        /// <param name="document">Parent document.</param>
        public WordDocumentStatistics(WordDocument document) {
            _document = document;
            document.Statistics = this;
        }

        /// <summary>
        /// Gets the total number of pages in the document.
        /// </summary>
        public int Pages {
            get {
                var pagesText = _document.ApplicationProperties?.Pages;
                if (!string.IsNullOrEmpty(pagesText) && int.TryParse(pagesText, out var pages)) {
                    return pages;
                }

                return _document.ParagraphsPageBreaks.Count + _document.Sections.Count;
            }
        }

        /// <summary>
        /// Gets the total number of paragraphs in the document.
        /// </summary>
        public int Paragraphs => _document.Paragraphs.Count;

        /// <summary>
        /// Gets the total number of words in the document.
        /// </summary>
        public int Words {
            get {
                var wordsProp = _document.ApplicationProperties?.Words;
                if (wordsProp != null && int.TryParse(wordsProp.Text, out var propertyWords)) {
                    return propertyWords;
                }

                int count = 0;
                foreach (var paragraph in _document.Paragraphs) {
                    var text = paragraph.Text;
                    if (!string.IsNullOrWhiteSpace(text)) {
                        count += text.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Length;
                    }
                }

                return count;
            }
        }

        /// <summary>
        /// Gets the total number of images in the document.
        /// </summary>
        public int Images => _document.Images.Count;

        /// <summary>
        /// Gets the total number of tables in the document.
        /// </summary>
        public int Tables => _document.TablesIncludingNestedTables.Count;

        /// <summary>
        /// Gets the total number of charts in the document.
        /// </summary>
        public int Charts => _document.Charts.Count;

        /// <summary>
        /// Gets the total number of shapes in the document.
        /// </summary>
        public int Shapes => _document.Shapes.Count;

        /// <summary>
        /// Gets the total number of bookmarks in the document.
        /// </summary>
        public int Bookmarks => _document.Bookmarks.Count;

        /// <summary>
        /// Gets the total number of lists in the document.
        /// </summary>
        public int Lists => _document.Lists.Count;

        /// <summary>
        /// Gets the total number of characters in the document (excluding spaces).
        /// </summary>
        public int Characters {
            get {
                var charText = _document.ApplicationProperties?.Characters;
                if (!string.IsNullOrEmpty(charText) && int.TryParse(charText, out var chars)) {
                    return chars;
                }

                return _document.Paragraphs.Sum(p => p.Text.Replace(" ", "").Length);
            }
        }

        /// <summary>
        /// Gets the total number of characters in the document including spaces.
        /// </summary>
        public int CharactersWithSpaces {
            get {
                var charText = _document.ApplicationProperties?.CharactersWithSpaces;
                if (!string.IsNullOrEmpty(charText) && int.TryParse(charText, out var chars)) {
                    return chars;
                }

                return _document.Paragraphs.Sum(p => p.Text.Length);
            }
        }
    }
}
