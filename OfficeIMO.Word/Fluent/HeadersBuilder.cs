using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for headers.
    /// </summary>
    public class HeadersBuilder {
        private readonly WordFluentDocument _fluent;

        internal HeadersBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        private WordHeader GetOrCreate(HeaderFooterValues type) {
            var document = _fluent.Document;
            var section = document.Sections[0];

            WordHeader? header;
            if (type == HeaderFooterValues.First) {
                header = document.Header?.First;
            } else if (type == HeaderFooterValues.Even) {
                header = document.Header?.Even;
            } else {
                header = document.Header?.Default;
            }

            if (header == null) {
                WordHeadersAndFooters.AddHeaderReference(document, section, type);
                if (type == HeaderFooterValues.First) {
                    header = document.Header!.First;
                } else if (type == HeaderFooterValues.Even) {
                    header = document.Header!.Even;
                } else {
                    header = document.Header!.Default;
                }
            }

            return header ?? throw new InvalidOperationException("Failed to create header instance.");
        }

        /// <summary>
        /// Adds content to the default header.
        /// </summary>
        public HeadersBuilder Default(Action<HeaderContentBuilder> action) {
            var header = GetOrCreate(HeaderFooterValues.Default);
            action(new HeaderContentBuilder(_fluent, header));
            return this;
        }

        /// <summary>
        /// Adds content to the first-page header.
        /// </summary>
        public HeadersBuilder First(Action<HeaderContentBuilder> action) {
            var header = GetOrCreate(HeaderFooterValues.First);
            action(new HeaderContentBuilder(_fluent, header));
            return this;
        }

        /// <summary>
        /// Adds content to the even-page header.
        /// </summary>
        public HeadersBuilder Even(Action<HeaderContentBuilder> action) {
            var header = GetOrCreate(HeaderFooterValues.Even);
            action(new HeaderContentBuilder(_fluent, header));
            return this;
        }

        /// <summary>
        /// Adds content to the odd-page header.
        /// </summary>
        public HeadersBuilder Odd(Action<HeaderContentBuilder> action) {
            return Default(action);
        }

        /// <summary>
        /// Adds a new header containing the specified text.
        /// </summary>
        public HeadersBuilder AddHeader(string text) {
            return Default(h => h.Paragraph(text));
        }
    }

    /// <summary>
    /// Allows adding content to a specific header.
    /// </summary>
    public class HeaderContentBuilder {
        private readonly WordFluentDocument _fluent;
        private readonly WordHeader _header;

        internal HeaderContentBuilder(WordFluentDocument fluent, WordHeader header) {
            _fluent = fluent;
            _header = header;
        }

        /// <summary>
        /// Adds a paragraph with the specified text.
        /// </summary>
        public HeaderContentBuilder Paragraph(string text) {
            _header.AddParagraph(text);
            return this;
        }

        /// <summary>
        /// Adds a paragraph configured using the supplied action.
        /// </summary>
        public HeaderContentBuilder Paragraph(Action<ParagraphBuilder> action) {
            var paragraph = _header.AddParagraph();
            action(new ParagraphBuilder(_fluent, paragraph));
            return this;
        }

        /// <summary>
        /// Adds an image to the header.
        /// </summary>
        public HeaderContentBuilder Image(string path, double? width = null, double? height = null, WrapTextImage wrapImage = WrapTextImage.InLineWithText, string description = "") {
            var paragraph = _header.AddParagraph();
            paragraph.AddImage(path, width, height, wrapImage, description);
            return this;
        }
        /// <summary>
        /// Adds a table to the header.
        /// </summary>
        public HeaderContentBuilder Table(int rows, int columns, WordTableStyle tableStyle = WordTableStyle.TableGrid, Action<WordTable>? configure = null) {
            var table = _header.AddTable(rows, columns, tableStyle);
            configure?.Invoke(table);
            return this;
        }
    }
}

