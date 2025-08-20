namespace OfficeIMO.Visio.Fluent {
    /// <summary>
    /// Provides a fluent wrapper for <see cref="VisioDocument"/> allowing chained configuration.
    /// </summary>
    public class VisioFluentDocument {
        private readonly VisioDocument _document;

        /// <summary>
        /// Initializes a new instance of the <see cref="VisioFluentDocument"/> class.
        /// </summary>
        /// <param name="document">The underlying <see cref="VisioDocument"/>.</param>
        public VisioFluentDocument(VisioDocument document) {
            _document = document;
        }

        /// <summary>
        /// Adds a page and returns the fluent document for chaining.
        /// </summary>
        /// <param name="name">Name of the page.</param>
        /// <param name="page">The created page.</param>
        public VisioFluentDocument AddPage(string name, out VisioPage page) {
            page = _document.AddPage(name);
            return this;
        }

        /// <summary>
        /// Ends fluent configuration and returns the underlying document.
        /// </summary>
        public VisioDocument End() {
            return _document;
        }
    }
}
