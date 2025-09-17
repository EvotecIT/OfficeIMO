using System;
using DocumentFormat.OpenXml.Bibliography;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a bibliographic source within a Word document.
    /// </summary>
    public class WordBibliographySource {
        internal Source Source { get; }

        private string? _tag;

        /// <summary>
        /// Initializes a new source with the given tag and type.
        /// </summary>
        public WordBibliographySource(string tag, DataSourceValues type) {
            Source = new Source();
            Tag = tag;
            SourceType = type;
        }

        internal WordBibliographySource(Source source) {
            Source = source;
            _tag = source.Tag?.Text;
        }

        /// <summary>
        /// Gets or sets the tag used to reference the source in citation fields.
        /// </summary>
        public string? Tag {
            get => _tag;
            set {
                _tag = value;
                if (string.IsNullOrEmpty(value)) {
                    Source.Tag = null;
                } else {
                    Source.Tag = new Tag { Text = value! };
                }
            }
        }

        /// <summary>
        /// Gets or sets the type of the source.
        /// </summary>
        public DataSourceValues SourceType {
            get {
                if (Enum.TryParse(Source.SourceType?.Text, out DataSourceValues val)) {
                    return val;
                }
                return DataSourceValues.Miscellaneous;
            }
            set {
                if (Source.SourceType == null) Source.SourceType = new SourceType();
                Source.SourceType.Text = value.ToString();
            }
        }

        /// <summary>
        /// Gets or sets the title of the source.
        /// </summary>
        public string? Title {
            get => Source.Title?.Text;
            set {
                if (Source.Title == null) Source.Title = new Title();
                Source.Title.Text = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets the corporate author of the source.
        /// </summary>
        public string? Author {
            get => Source.AuthorList?.GetFirstChild<Author>()?.GetFirstChild<Corporate>()?.Text;
            set {
                if (Source.AuthorList == null) Source.AuthorList = new AuthorList();
                var author = Source.AuthorList.GetFirstChild<Author>();
                if (author == null) {
                    author = new Author();
                    Source.AuthorList.Append(author);
                }
                var corp = author.GetFirstChild<Corporate>();
                if (corp == null) {
                    corp = new Corporate();
                    author.Append(corp);
                }
                corp.Text = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets the publication year of the source.
        /// </summary>
        public string? Year {
            get => Source.Year?.Text;
            set {
                if (Source.Year == null) Source.Year = new Year();
                Source.Year.Text = value ?? string.Empty;
            }
        }

        internal Source ToOpenXml() {
            // ensure Tag element reflects current state
            if (string.IsNullOrEmpty(_tag)) {
                Source.Tag = null;
            } else {
                Source.Tag = new Tag { Text = _tag! };
            }
            return Source;
        }
    }
}
