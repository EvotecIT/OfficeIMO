using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Built-in (core) document properties for an Excel workbook, such as Title and Author.
    /// The wrapper remains associated with its workbook across saves.
    /// </summary>
    public sealed class ExcelBuiltinDocumentProperties {
        private SpreadsheetDocument Document => _document._spreadSheetDocument
            ?? throw new ObjectDisposedException(nameof(ExcelDocument));
        private readonly ExcelDocument _document;

        /// <summary>
        /// Creates a new wrapper for core document properties bound to the given document.
        /// </summary>
        public ExcelBuiltinDocumentProperties(ExcelDocument document) {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            document.BuiltinDocumentProperties = this;
            EnsureCorePropertiesPart();
        }

        private void EnsureCorePropertiesPart() {
            // Touch the package properties to ensure the backing object is initialized.
            try {
                GC.KeepAlive(Document.PackageProperties);
            } catch (ObjectDisposedException) {
                // Package properties are best-effort when the package is already disposed.
            } catch (InvalidOperationException) {
                // Some package states do not expose core properties yet; callers retry on set.
            }
        }

        /// <summary>The document creator/author.</summary>
        public string? Creator {
            get { return Document.PackageProperties.Creator; }
            set { EnsureCorePropertiesPart(); Document.PackageProperties.Creator = value; _document.MarkPackagePropertiesDirty(); }
        }

        /// <summary>Document title.</summary>
        public string? Title {
            get { return Document.PackageProperties.Title; }
            set { EnsureCorePropertiesPart(); Document.PackageProperties.Title = value; _document.MarkPackagePropertiesDirty(); }
        }

        /// <summary>Document description/summary.</summary>
        public string? Description {
            get { return Document.PackageProperties.Description; }
            set { EnsureCorePropertiesPart(); Document.PackageProperties.Description = value; _document.MarkPackagePropertiesDirty(); }
        }

        /// <summary>Document category.</summary>
        public string? Category {
            get { return Document.PackageProperties.Category; }
            set { EnsureCorePropertiesPart(); Document.PackageProperties.Category = value; _document.MarkPackagePropertiesDirty(); }
        }

        /// <summary>Comma- or space-separated keywords.</summary>
        public string? Keywords {
            get { return Document.PackageProperties.Keywords; }
            set { EnsureCorePropertiesPart(); Document.PackageProperties.Keywords = value; _document.MarkPackagePropertiesDirty(); }
        }

        /// <summary>Document subject.</summary>
        public string? Subject {
            get { return Document.PackageProperties.Subject; }
            set { EnsureCorePropertiesPart(); Document.PackageProperties.Subject = value; _document.MarkPackagePropertiesDirty(); }
        }

        /// <summary>Revision number.</summary>
        public string? Revision {
            get { return Document.PackageProperties.Revision; }
            set { EnsureCorePropertiesPart(); Document.PackageProperties.Revision = value; _document.MarkPackagePropertiesDirty(); }
        }

        /// <summary>Last modified by user.</summary>
        public string? LastModifiedBy {
            get { return Document.PackageProperties.LastModifiedBy; }
            set { EnsureCorePropertiesPart(); Document.PackageProperties.LastModifiedBy = value; _document.MarkPackagePropertiesDirty(); }
        }

        /// <summary>Document version.</summary>
        public string? Version {
            get { return Document.PackageProperties.Version; }
            set { EnsureCorePropertiesPart(); Document.PackageProperties.Version = value; _document.MarkPackagePropertiesDirty(); }
        }

        /// <summary>Creation timestamp (UTC).</summary>
        public DateTime? Created {
            get { return Document.PackageProperties.Created; }
            set { EnsureCorePropertiesPart(); Document.PackageProperties.Created = value; _document.MarkPackagePropertiesDirty(); }
        }

        /// <summary>Last modified timestamp (UTC).</summary>
        public DateTime? Modified {
            get { return Document.PackageProperties.Modified; }
            set { EnsureCorePropertiesPart(); Document.PackageProperties.Modified = value; _document.MarkPackagePropertiesDirty(); }
        }

        /// <summary>Last printed timestamp (UTC).</summary>
        public DateTime? LastPrinted {
            get { return Document.PackageProperties.LastPrinted; }
            set { EnsureCorePropertiesPart(); Document.PackageProperties.LastPrinted = value; _document.MarkPackagePropertiesDirty(); }
        }
    }
}
