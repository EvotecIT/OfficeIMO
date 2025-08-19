using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

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
                    footer = document.Footer.First;
                } else if (type == HeaderFooterValues.Even) {
                    footer = document.Footer.Even;
                } else {
                    footer = document.Footer.Default;
                }
            }

            return footer;
        }

        public FootersBuilder Default(Action<FooterContentBuilder> action) {
            var footer = GetOrCreate(HeaderFooterValues.Default);
            action(new FooterContentBuilder(_fluent, footer));
            return this;
        }

        public FootersBuilder First(Action<FooterContentBuilder> action) {
            var footer = GetOrCreate(HeaderFooterValues.First);
            action(new FooterContentBuilder(_fluent, footer));
            return this;
        }

        public FootersBuilder Even(Action<FooterContentBuilder> action) {
            var footer = GetOrCreate(HeaderFooterValues.Even);
            action(new FooterContentBuilder(_fluent, footer));
            return this;
        }

        public FootersBuilder Odd(Action<FooterContentBuilder> action) {
            return Default(action);
        }

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

        public FooterContentBuilder Paragraph(string text) {
            _footer.AddParagraph(text);
            return this;
        }

        public FooterContentBuilder Paragraph(Action<ParagraphBuilder> action) {
            var paragraph = _footer.AddParagraph();
            action(new ParagraphBuilder(_fluent, paragraph));
            return this;
        }

        public FooterContentBuilder Image(string path, double? width = null, double? height = null, WrapTextImage wrapImage = WrapTextImage.InLineWithText, string description = "") {
            var paragraph = _footer.AddParagraph();
            paragraph.AddImage(path, width, height, wrapImage, description);
            return this;
        }

        public FooterContentBuilder Table(int rows, int columns, WordTableStyle tableStyle = WordTableStyle.TableGrid, Action<WordTable>? configure = null) {
            var table = _footer.AddTable(rows, columns, tableStyle);
            configure?.Invoke(table);
            return this;
        }
    }
}

