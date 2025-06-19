using System;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word {
    public class BuiltinDocumentProperties {
        private WordprocessingDocument _wordprocessingDocument;
        private WordDocument _document;

        public BuiltinDocumentProperties(WordDocument document) {
            _document = document;
            _wordprocessingDocument = document._wordprocessingDocument;

            document.BuiltinDocumentProperties = this;
        }
        /// <summary>
        /// Gets or sets the author of the document.
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
        /// Gets or sets the description or abstract for the document.
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
        /// Gets or sets the category of the document.
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
        /// Gets or sets keywords that help with searching and indexing the document.
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
        /// Gets or sets the subject of the document.
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
        /// Gets or sets the revision number of the document.
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
        /// Gets or sets the name of the person who last modified the document.
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
        /// Gets or sets the version of the document.
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
        /// Gets or sets the document creation date and time.
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
        /// Gets or sets the date and time the document was last modified.
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
        /// Gets or sets the date the document was last printed.
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
