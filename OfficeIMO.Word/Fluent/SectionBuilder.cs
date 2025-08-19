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
    }
}
