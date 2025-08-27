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

        /// <summary>
        /// Gets the section being configured.
        /// </summary>
        public WordSection? Section => _section;

        /// <summary>
        /// Starts a new section on the next page.
        /// </summary>
        public SectionBuilder New() {
            return New(SectionMarkValues.NextPage);
        }

        /// <summary>
        /// Starts a new section using the specified break type.
        /// </summary>
        /// <param name="breakType">Section break type.</param>
        public SectionBuilder New(SectionMarkValues breakType) {
            var section = _fluent.Document.AddSection(breakType);
            return new SectionBuilder(_fluent, section);
        }

        /// <summary>
        /// Enables page numbering for the section.
        /// </summary>
        /// <param name="restart">Restart numbering at 1.</param>
        public SectionBuilder PageNumbering(bool restart = false) {
            if (_section == null) {
                throw new InvalidOperationException("No section available to configure.");
            }

            _section.AddPageNumbering(restart ? 1 : (int?)null);
            return this;
        }

        /// <summary>
        /// Enables page numbering with a specific number format.
        /// </summary>
        /// <param name="format">Number format.</param>
        /// <param name="restart">Restart numbering at 1.</param>
        public SectionBuilder PageNumbering(NumberFormatValues format, bool restart = false) {
            if (_section == null) {
                throw new InvalidOperationException("No section available to configure.");
            }

            _section.AddPageNumbering(restart ? 1 : (int?)null, format);
            return this;
        }

        /// <summary>
        /// Sets the number of columns for the section.
        /// </summary>
        /// <param name="count">Column count.</param>
        public SectionBuilder Columns(int count) {
            if (_section == null) {
                throw new InvalidOperationException("No section available to configure.");
            }

            _section.ColumnCount = count;
            return this;
        }

        /// <summary>
        /// Sets the margins for the section.
        /// </summary>
        /// <param name="margins">Margin values.</param>
        public SectionBuilder Margins(WordMargin margins) {
            if (_section == null) {
                throw new InvalidOperationException("No section available to configure.");
            }

            _section.SetMargins(margins);
            return this;
        }

        /// <summary>
        /// Sets the page size for the section.
        /// </summary>
        /// <param name="pageSize">Page size definition.</param>
        public SectionBuilder Size(WordPageSize pageSize) {
            if (_section == null) {
                throw new InvalidOperationException("No section available to configure.");
            }

            _section.PageSettings.PageSize = pageSize;
            return this;
        }

        /// <summary>
        /// Adds a paragraph to the section.
        /// </summary>
        /// <param name="action">Configuration action for the paragraph.</param>
        public SectionBuilder Paragraph(Action<ParagraphBuilder> action) {
            var paragraph = _fluent.Document.AddParagraph();
            action(new ParagraphBuilder(_fluent, paragraph));
            return this;
        }

        /// <summary>
        /// Adds a table to the section.
        /// </summary>
        /// <param name="action">Configuration action for the table.</param>
        public SectionBuilder Table(Action<TableBuilder> action) {
            action(new TableBuilder(_fluent));
            return this;
        }
    }
}
