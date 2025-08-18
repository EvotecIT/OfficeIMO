namespace OfficeIMO.Visio.Fluent {
    /// <summary>
    /// Provides fluent helpers for <see cref="VisioDocument"/>.
    /// </summary>
    public static class VisioDocumentFluent {
        /// <summary>
        /// Adds a page and returns the document for chaining.
        /// </summary>
        public static VisioDocument AddPage(this VisioDocument document, string name, out VisioPage page) {
            page = document.AddPage(name);
            return document;
        }
    }
}

