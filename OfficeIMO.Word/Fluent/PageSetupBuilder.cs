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

        /// <summary>
        /// Sets the page orientation.
        /// </summary>
        /// <param name="orientation">Orientation value.</param>
        public PageSetupBuilder Orientation(PageOrientationValues orientation) {
            _fluent.Document.PageOrientation = orientation;
            return this;
        }

        /// <summary>
        /// Sets the page size.
        /// </summary>
        /// <param name="pageSize">Page size definition.</param>
        public PageSetupBuilder Size(WordPageSize pageSize) {
            _fluent.Document.PageSettings.PageSize = pageSize;
            return this;
        }

        /// <summary>
        /// Sets page margins for the first section.
        /// </summary>
        /// <param name="margin">Margin values.</param>
        public PageSetupBuilder Margins(WordMargin margin) {
            _fluent.Document.Sections[0].SetMargins(margin);
            return this;
        }

        /// <summary>
        /// Configures whether the first page uses a different header and footer.
        /// </summary>
        /// <param name="value">True to enable different first page.</param>
        public PageSetupBuilder DifferentFirstPage(bool value = true) {
            _fluent.Document.DifferentFirstPage = value;
            return this;
        }

        /// <summary>
        /// Configures whether odd and even pages have different headers and footers.
        /// </summary>
        /// <param name="value">True to enable different headers and footers.</param>
        public PageSetupBuilder DifferentOddAndEvenPages(bool value = true) {
            _fluent.Document.DifferentOddAndEvenPages = value;
            return this;
        }
    }
}
