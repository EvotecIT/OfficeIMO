namespace OfficeIMO.Visio.Fluent {
    /// <summary>
    /// Extension methods for <see cref="VisioDocument"/> to enable fluent configuration.
    /// </summary>
    public static class VisioFluentDocumentExtensions {
        /// <summary>
        /// Wraps the document in a <see cref="VisioFluentDocument"/> for fluent configuration.
        /// </summary>
        /// <param name="doc">The document to wrap.</param>
        public static VisioFluentDocument AsFluent(this VisioDocument doc) {
            return new VisioFluentDocument(doc);
        }
    }
}
