using DocumentFormat.OpenXml.ExtendedProperties;

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
        public InfoBuilder Title(string title) {
            _fluent.Document.BuiltinDocumentProperties.Title = title;
            return this;
        }

        /// <summary>
        /// Sets the document author.
        /// </summary>
        /// <param name="author">Author name.</param>
        public InfoBuilder Author(string author) {
            _fluent.Document.BuiltinDocumentProperties.Creator = author;
            return this;
        }

        /// <summary>
        /// Sets the document subject.
        /// </summary>
        /// <param name="subject">Subject text.</param>
        public InfoBuilder Subject(string subject) {
            _fluent.Document.BuiltinDocumentProperties.Subject = subject;
            return this;
        }

        /// <summary>
        /// Sets the document keywords.
        /// </summary>
        /// <param name="keywords">Keywords list.</param>
        public InfoBuilder Keywords(string keywords) {
            _fluent.Document.BuiltinDocumentProperties.Keywords = keywords;
            return this;
        }

        /// <summary>
        /// Sets the document comments.
        /// </summary>
        /// <param name="comments">Comments text.</param>
        public InfoBuilder Comments(string comments) {
            _fluent.Document.BuiltinDocumentProperties.Description = comments;
            return this;
        }

        /// <summary>
        /// Sets the document category.
        /// </summary>
        /// <param name="category">Category text.</param>
        public InfoBuilder Category(string category) {
            _fluent.Document.BuiltinDocumentProperties.Category = category;
            return this;
        }

        /// <summary>
        /// Sets the company name.
        /// </summary>
        /// <param name="company">Company name.</param>
        public InfoBuilder Company(string company) {
            _fluent.Document.ApplicationProperties.Company = company;
            return this;
        }

        /// <summary>
        /// Sets the manager name.
        /// </summary>
        /// <param name="manager">Manager name.</param>
        public InfoBuilder Manager(string manager) {
            _fluent.Document.ApplicationProperties.Manager = new Manager { Text = manager };
            return this;
        }

        /// <summary>
        /// Sets the user who last modified the document.
        /// </summary>
        /// <param name="lastModifiedBy">Last modified by.</param>
        public InfoBuilder LastModifiedBy(string lastModifiedBy) {
            _fluent.Document.BuiltinDocumentProperties.LastModifiedBy = lastModifiedBy;
            return this;
        }

        /// <summary>
        /// Sets the document revision.
        /// </summary>
        /// <param name="revision">Revision value.</param>
        public InfoBuilder Revision(string revision) {
            _fluent.Document.BuiltinDocumentProperties.Revision = revision;
            return this;
        }

        /// <summary>
        /// Adds or updates a custom document property.
        /// </summary>
        /// <param name="name">Property name.</param>
        /// <param name="value">Property value.</param>
        public InfoBuilder Custom(string name, object value) {
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
