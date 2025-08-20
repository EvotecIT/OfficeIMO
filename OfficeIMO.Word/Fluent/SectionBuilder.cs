using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for sections.
    /// </summary>
    public class SectionBuilder {
        private readonly WordFluentDocument _fluent;
        private readonly WordSection? _section;

        internal SectionBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        internal SectionBuilder(WordFluentDocument fluent, WordSection section) {
            _fluent = fluent;
            _section = section;
        }

        public WordSection? Section => _section;

        public SectionBuilder New() {
            return New(SectionMarkValues.NextPage);
        }

        public SectionBuilder New(SectionMarkValues breakType) {
            var section = _fluent.Document.AddSection(breakType);
            return new SectionBuilder(_fluent, section);
        }

        public SectionBuilder PageNumbering(bool restart = false) {
            if (_section == null) {
                throw new InvalidOperationException("No section available to configure.");
            }

            _section.AddPageNumbering(restart ? 1 : (int?)null);
            return this;
        }

        public SectionBuilder PageNumbering(NumberFormatValues format, bool restart = false) {
            if (_section == null) {
                throw new InvalidOperationException("No section available to configure.");
            }

            _section.AddPageNumbering(restart ? 1 : (int?)null, format);
            return this;
        }

        public SectionBuilder Columns(int count) {
            if (_section == null) {
                throw new InvalidOperationException("No section available to configure.");
            }

            _section.ColumnCount = count;
            return this;
        }

        public SectionBuilder Margins(WordMargin margins) {
            if (_section == null) {
                throw new InvalidOperationException("No section available to configure.");
            }

            _section.SetMargins(margins);
            return this;
        }

        public SectionBuilder Size(WordPageSize pageSize) {
            if (_section == null) {
                throw new InvalidOperationException("No section available to configure.");
            }

            _section.PageSettings.PageSize = pageSize;
            return this;
        }

        public SectionBuilder Paragraph(Action<ParagraphBuilder> action) {
            var paragraph = _fluent.Document.AddParagraph();
            action(new ParagraphBuilder(_fluent, paragraph));
            return this;
        }

        public SectionBuilder Table(Action<TableBuilder> action) {
            action(new TableBuilder(_fluent));
            return this;
        }
    }
}
