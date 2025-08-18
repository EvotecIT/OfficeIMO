using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for page settings.
    /// </summary>
    public class PageBuilder {
        private readonly WordFluentDocument _fluent;

        internal PageBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        public WordFluentDocument SetOrientation(PageOrientationValues orientation) {
            _fluent.Document.PageOrientation = orientation;
            return _fluent;
        }
    }
}
