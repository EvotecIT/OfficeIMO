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

        public SectionBuilder AddSection(SectionMarkValues? mark = null) {
            var section = _fluent.Document.AddSection(mark);
            return new SectionBuilder(_fluent, section);
        }

        public SectionBuilder SetSectionBreak(SectionMarkValues mark) {
            if (_section == null) {
                throw new InvalidOperationException("No section available to configure.");
            }

            var sectionType = _section._sectionProperties.GetFirstChild<SectionType>();
            if (sectionType == null) {
                sectionType = new SectionType();
                _section._sectionProperties.Append(sectionType);
            }

            sectionType.Val = mark;
            return this;
        }

        public SectionBuilder SetPageNumbering(NumberFormatValues? format = null, bool restart = false, int startNumber = 1) {
            if (_section == null) {
                throw new InvalidOperationException("No section available to configure.");
            }

            _section.AddPageNumbering(restart ? startNumber : (int?)null, format);
            return this;
        }

        public SectionBuilder SetMargins(WordMargin margins) {
            if (_section == null) {
                throw new InvalidOperationException("No section available to configure.");
            }

            _section.SetMargins(margins);
            return this;
        }

        public SectionBuilder SetPageSize(WordPageSize pageSize) {
            if (_section == null) {
                throw new InvalidOperationException("No section available to configure.");
            }

            _section.PageSettings.PageSize = pageSize;
            return this;
        }
    }
}
