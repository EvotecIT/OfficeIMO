using System;

namespace OfficeIMO.Visio.Fluent {
    /// <summary>
    /// Builder for document information properties.
    /// </summary>
    public class InfoBuilder {
        private readonly VisioFluentDocument _fluent;

        internal InfoBuilder(VisioFluentDocument fluent) {
            _fluent = fluent;
        }

        /// <summary>
        /// Sets the document title.
        /// </summary>
        /// <param name="title">Title to assign.</param>
        public InfoBuilder Title(string title) {
            _fluent.Document.Title = title;
            return this;
        }

        /// <summary>
        /// Sets the document author.
        /// </summary>
        /// <param name="author">Author name.</param>
        public InfoBuilder Author(string author) {
            _fluent.Document.Author = author;
            return this;
        }
    }
}

