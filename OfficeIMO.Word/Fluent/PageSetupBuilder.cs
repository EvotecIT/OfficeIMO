using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for page setup settings.
    /// </summary>
    public class PageSetupBuilder {
        private readonly WordFluentDocument _fluent;

        internal PageSetupBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        public PageSetupBuilder Orientation(PageOrientationValues orientation) {
            _fluent.Document.PageOrientation = orientation;
            return this;
        }

        public PageSetupBuilder Size(WordPageSize pageSize) {
            _fluent.Document.PageSettings.PageSize = pageSize;
            return this;
        }

        public PageSetupBuilder Margins(WordMargin margin) {
            _fluent.Document.Sections[0].SetMargins(margin);
            return this;
        }

        public PageSetupBuilder DifferentFirstPage(bool value = true) {
            _fluent.Document.DifferentFirstPage = value;
            return this;
        }

        public PageSetupBuilder DifferentOddAndEvenPages(bool value = true) {
            _fluent.Document.DifferentOddAndEvenPages = value;
            return this;
        }
    }
}
