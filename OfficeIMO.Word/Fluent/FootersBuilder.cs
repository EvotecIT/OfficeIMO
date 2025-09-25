using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for footers.
    /// </summary>
    public class FootersBuilder {
        private readonly WordFluentDocument _fluent;

        internal FootersBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        private WordFooter GetOrCreate(HeaderFooterValues type) {
            var document = _fluent.Document;
            var section = document.Sections[0];

            WordFooter? footer;
            if (type == HeaderFooterValues.First) {
                footer = document.Footer?.First;
            } else if (type == HeaderFooterValues.Even) {
                footer = document.Footer?.Even;
            } else {
                footer = document.Footer?.Default;
            }

            if (footer == null) {
                WordHeadersAndFooters.AddFooterReference(document, section, type);
                if (type == HeaderFooterValues.First) {
                    footer = document.Footer!.First;
                } else if (type == HeaderFooterValues.Even) {
                    footer = document.Footer!.Even;
                } else {
                    footer = document.Footer!.Default;
                }
            }

            return footer ?? throw new InvalidOperationException("Failed to create footer instance.");
        }

        /// <summary>
        /// Adds content to the default footer.
        /// </summary>
        public FootersBuilder Default(Action<FooterContentBuilder> action) {
            var footer = GetOrCreate(HeaderFooterValues.Default);
            action(new FooterContentBuilder(_fluent, footer));
            return this;
        }

        /// <summary>
        /// Adds content to the first-page footer.
        /// </summary>
        public FootersBuilder First(Action<FooterContentBuilder> action) {
            var footer = GetOrCreate(HeaderFooterValues.First);
            action(new FooterContentBuilder(_fluent, footer));
            return this;
        }

        /// <summary>
        /// Adds content to the even-page footer.
        /// </summary>
        public FootersBuilder Even(Action<FooterContentBuilder> action) {
            var footer = GetOrCreate(HeaderFooterValues.Even);
            action(new FooterContentBuilder(_fluent, footer));
            return this;
        }

        /// <summary>
        /// Adds content to the odd-page footer.
        /// </summary>
        public FootersBuilder Odd(Action<FooterContentBuilder> action) {
            return Default(action);
        }

        /// <summary>
        /// Adds a new footer containing the specified text.
        /// </summary>
        public FootersBuilder AddFooter(string text) {
            return Default(f => f.Paragraph(text));
        }
    }

    /// <summary>
    /// Allows adding content to a specific footer.
    /// </summary>
    public class FooterContentBuilder {
        private readonly WordFluentDocument _fluent;
        private readonly WordFooter _footer;

        internal FooterContentBuilder(WordFluentDocument fluent, WordFooter footer) {
            _fluent = fluent;
            _footer = footer;
        }

        /// <summary>
        /// Adds a paragraph with the specified text.
        /// </summary>
        public FooterContentBuilder Paragraph(string text) {
            _footer.AddParagraph(text);
            return this;
        }

        /// <summary>
        /// Adds a paragraph configured using the supplied action.
        /// </summary>
        public FooterContentBuilder Paragraph(Action<ParagraphBuilder> action) {
            var paragraph = _footer.AddParagraph();
            action(new ParagraphBuilder(_fluent, paragraph));
            return this;
        }

        /// <summary>
        /// Adds an image to the footer.
        /// </summary>
        public FooterContentBuilder Image(string path, double? width = null, double? height = null, WrapTextImage wrapImage = WrapTextImage.InLineWithText, string description = "") {
            var paragraph = _footer.AddParagraph();
            paragraph.AddImage(path, width, height, wrapImage, description);
            return this;
        }

        /// <summary>
        /// Adds a table to the footer.
        /// </summary>
        public FooterContentBuilder Table(int rows, int columns, WordTableStyle tableStyle = WordTableStyle.TableGrid, Action<WordTable>? configure = null) {
            var table = _footer.AddTable(rows, columns, tableStyle);
            configure?.Invoke(table);
            return this;
        }
    }
}

