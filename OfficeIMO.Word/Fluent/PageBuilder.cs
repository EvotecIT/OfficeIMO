using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for page settings.
    /// </summary>
    public class PageBuilder {
        private readonly WordFluentDocument _fluent;

        internal PageBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        public PageBuilder SetOrientation(PageOrientationValues orientation) {
            _fluent.Document.PageOrientation = orientation;
            return this;
        }

        public PageBuilder SetPaperSize(WordPageSize pageSize) {
            _fluent.Document.PageSettings.PageSize = pageSize;
            return this;
        }

        public PageBuilder SetMargins(WordMargin margin) {
            _fluent.Document.Sections[0].SetMargins(margin);
            return this;
        }

        public PageBuilder DifferentFirstPage(bool value = true) {
            _fluent.Document.DifferentFirstPage = value;
            return this;
        }

        public PageBuilder DifferentOddAndEvenPages(bool value = true) {
            _fluent.Document.DifferentOddAndEvenPages = value;
            return this;
        }
    }
}
