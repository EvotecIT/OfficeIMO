using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for sections.
    /// </summary>
    public class SectionBuilder {
        private readonly WordFluentDocument _fluent;

        internal SectionBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        public WordFluentDocument AddSection(SectionMarkValues? mark = null) {
            _fluent.Document.AddSection(mark);
            return _fluent;
        }
    }
}
