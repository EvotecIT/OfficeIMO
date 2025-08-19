using System;

namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for document information such as properties.
    /// </summary>
    public class InfoBuilder {
        private readonly WordFluentDocument _fluent;

        internal InfoBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        /// <summary>
        /// Sets the document title.
        /// </summary>
        /// <param name="title">Title to assign.</param>
        public InfoBuilder SetTitle(string title) {
            _fluent.Document.BuiltinDocumentProperties.Title = title;
            return this;
        }

        /// <summary>
        /// Sets the document author.
        /// </summary>
        /// <param name="author">Author name.</param>
        public InfoBuilder SetAuthor(string author) {
            _fluent.Document.BuiltinDocumentProperties.Creator = author;
            return this;
        }

        /// <summary>
        /// Sets the document subject.
        /// </summary>
        /// <param name="subject">Subject text.</param>
        public InfoBuilder SetSubject(string subject) {
            _fluent.Document.BuiltinDocumentProperties.Subject = subject;
            return this;
        }

        /// <summary>
        /// Sets the document keywords.
        /// </summary>
        /// <param name="keywords">Keywords list.</param>
        public InfoBuilder SetKeywords(string keywords) {
            _fluent.Document.BuiltinDocumentProperties.Keywords = keywords;
            return this;
        }

        /// <summary>
        /// Sets the document comments.
        /// </summary>
        /// <param name="comments">Comments text.</param>
        public InfoBuilder SetComments(string comments) {
            _fluent.Document.BuiltinDocumentProperties.Description = comments;
            return this;
        }

        /// <summary>
        /// Adds or updates a custom document property.
        /// </summary>
        /// <param name="name">Property name.</param>
        /// <param name="value">Property value.</param>
        public InfoBuilder SetCustomProperty(string name, object value) {
            var property = value switch {
                bool b => new WordCustomProperty(b),
                DateTime dt => new WordCustomProperty(dt),
                double d => new WordCustomProperty(d),
                int i => new WordCustomProperty(i),
                string s => new WordCustomProperty(s),
                _ => new WordCustomProperty(value.ToString() ?? string.Empty)
            };

            _fluent.Document.CustomDocumentProperties[name] = property;
            return this;
        }
    }
}
