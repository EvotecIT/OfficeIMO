using System;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides strongly typed access to the built-in package properties.
    /// </summary>
    public class BuiltinDocumentProperties {
        private WordprocessingDocument _wordprocessingDocument;
        private WordDocument _document;

        /// <summary>
        /// Initializes a new instance bound to the specified document.
        /// </summary>
        /// <param name="document">Parent Word document.</param>
        public BuiltinDocumentProperties(WordDocument document) {
            _document = document;
            _wordprocessingDocument = document._wordprocessingDocument;

            document.BuiltinDocumentProperties = this;
        }
        /// <summary>
        /// Gets or sets the name of the document creator.
        /// </summary>
        public string Creator {
            get {
                return _wordprocessingDocument.PackageProperties.Creator;
            }
            set {
                _wordprocessingDocument.PackageProperties.Creator = value;
            }
        }
        /// <summary>
        /// Gets or sets the document title.
        /// </summary>
        public string Title {
            get {
                return _wordprocessingDocument.PackageProperties.Title;
            }
            set {
                _wordprocessingDocument.PackageProperties.Title = value;
            }
        }
        /// <summary>
        /// Gets or sets the document description.
        /// </summary>
        public string Description {
            get {
                return _wordprocessingDocument.PackageProperties.Description;
            }
            set {
                _wordprocessingDocument.PackageProperties.Description = value;
            }
        }
        /// <summary>
        /// Gets or sets the document category.
        /// </summary>
        public string Category {
            get {
                return _wordprocessingDocument.PackageProperties.Category;
            }
            set {
                _wordprocessingDocument.PackageProperties.Category = value;
            }
        }
        /// <summary>
        /// A delimited set of keywords (tags) to support searching and indexing the Package and content.
        /// </summary>
        public string Keywords {
            get {
                return _wordprocessingDocument.PackageProperties.Keywords;
            }
            set {
                _wordprocessingDocument.PackageProperties.Keywords = value;
            }
        }
        /// <summary>
        /// Gets or sets the document subject.
        /// </summary>
        public string Subject {
            get {
                return _wordprocessingDocument.PackageProperties.Subject;
            }
            set {
                _wordprocessingDocument.PackageProperties.Subject = value;
            }
        }
        /// <summary>
        /// Gets or sets the revision identifier.
        /// </summary>
        public string Revision {
            get {
                return _wordprocessingDocument.PackageProperties.Revision;
            }
            set {
                _wordprocessingDocument.PackageProperties.Revision = value;
            }
        }
        /// <summary>
        /// Gets or sets the user that last modified the document.
        /// </summary>
        public string LastModifiedBy {
            get {
                return _wordprocessingDocument.PackageProperties.LastModifiedBy;
            }
            set {
                _wordprocessingDocument.PackageProperties.LastModifiedBy = value;
            }
        }
        /// <summary>
        /// Gets or sets the document version.
        /// </summary>
        public string Version {
            get {
                return _wordprocessingDocument.PackageProperties.Version;
            }
            set {
                _wordprocessingDocument.PackageProperties.Version = value;
            }
        }
        /// <summary>
        /// Gets or sets the document creation time.
        /// </summary>
        public DateTime? Created {
            get {
                return _wordprocessingDocument.PackageProperties.Created;
            }
            set {
                _wordprocessingDocument.PackageProperties.Created = value;
            }
        }
        /// <summary>
        /// Gets or sets the last modified time of the document.
        /// </summary>
        public DateTime? Modified {
            get {
                return _wordprocessingDocument.PackageProperties.Modified;
            }
            set {
                _wordprocessingDocument.PackageProperties.Modified = value;
            }
        }
        /// <summary>
        /// Gets or sets the time the document was last printed.
        /// </summary>
        public DateTime? LastPrinted {
            get {
                return _wordprocessingDocument.PackageProperties.LastPrinted;
            }
            set {
                _wordprocessingDocument.PackageProperties.LastPrinted = value;
            }
        }
    }
}
