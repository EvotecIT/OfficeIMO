using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Built-in (core) document properties for an Excel workbook, such as Title and Author.
    /// </summary>
    public sealed class BuiltinDocumentProperties {
        private readonly SpreadsheetDocument _spreadsheetDocument;
        private readonly ExcelDocument _document;

        /// <summary>
        /// Creates a new wrapper for core document properties bound to the given document.
        /// </summary>
        public BuiltinDocumentProperties(ExcelDocument document) {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _spreadsheetDocument = document._spreadSheetDocument;
            document.BuiltinDocumentProperties = this;
            EnsureCorePropertiesPart();
        }

        private void EnsureCorePropertiesPart() {
            // Touch the package properties to ensure the backing object is initialized.
            try {
                GC.KeepAlive(_spreadsheetDocument.PackageProperties);
            } catch (ObjectDisposedException) {
                // Package properties are best-effort when the package is already disposed.
            } catch (InvalidOperationException) {
                // Some package states do not expose core properties yet; callers retry on set.
            }
        }

        /// <summary>The document creator/author.</summary>
        public string? Creator {
            get { return _spreadsheetDocument.PackageProperties.Creator; }
            set { EnsureCorePropertiesPart(); _spreadsheetDocument.PackageProperties.Creator = value; _document.MarkPackagePropertiesDirty(); }
        }

        /// <summary>Document title.</summary>
        public string? Title {
            get { return _spreadsheetDocument.PackageProperties.Title; }
            set { EnsureCorePropertiesPart(); _spreadsheetDocument.PackageProperties.Title = value; _document.MarkPackagePropertiesDirty(); }
        }

        /// <summary>Document description/summary.</summary>
        public string? Description {
            get { return _spreadsheetDocument.PackageProperties.Description; }
            set { EnsureCorePropertiesPart(); _spreadsheetDocument.PackageProperties.Description = value; _document.MarkPackagePropertiesDirty(); }
        }

        /// <summary>Document category.</summary>
        public string? Category {
            get { return _spreadsheetDocument.PackageProperties.Category; }
            set { EnsureCorePropertiesPart(); _spreadsheetDocument.PackageProperties.Category = value; _document.MarkPackagePropertiesDirty(); }
        }

        /// <summary>Comma- or space-separated keywords.</summary>
        public string? Keywords {
            get { return _spreadsheetDocument.PackageProperties.Keywords; }
            set { EnsureCorePropertiesPart(); _spreadsheetDocument.PackageProperties.Keywords = value; _document.MarkPackagePropertiesDirty(); }
        }

        /// <summary>Document subject.</summary>
        public string? Subject {
            get { return _spreadsheetDocument.PackageProperties.Subject; }
            set { EnsureCorePropertiesPart(); _spreadsheetDocument.PackageProperties.Subject = value; _document.MarkPackagePropertiesDirty(); }
        }

        /// <summary>Revision number.</summary>
        public string? Revision {
            get { return _spreadsheetDocument.PackageProperties.Revision; }
            set { EnsureCorePropertiesPart(); _spreadsheetDocument.PackageProperties.Revision = value; _document.MarkPackagePropertiesDirty(); }
        }

        /// <summary>Last modified by user.</summary>
        public string? LastModifiedBy {
            get { return _spreadsheetDocument.PackageProperties.LastModifiedBy; }
            set { EnsureCorePropertiesPart(); _spreadsheetDocument.PackageProperties.LastModifiedBy = value; _document.MarkPackagePropertiesDirty(); }
        }

        /// <summary>Document version.</summary>
        public string? Version {
            get { return _spreadsheetDocument.PackageProperties.Version; }
            set { EnsureCorePropertiesPart(); _spreadsheetDocument.PackageProperties.Version = value; _document.MarkPackagePropertiesDirty(); }
        }

        /// <summary>Creation timestamp (UTC).</summary>
        public DateTime? Created {
            get { return _spreadsheetDocument.PackageProperties.Created; }
            set { EnsureCorePropertiesPart(); _spreadsheetDocument.PackageProperties.Created = value; _document.MarkPackagePropertiesDirty(); }
        }

        /// <summary>Last modified timestamp (UTC).</summary>
        public DateTime? Modified {
            get { return _spreadsheetDocument.PackageProperties.Modified; }
            set { EnsureCorePropertiesPart(); _spreadsheetDocument.PackageProperties.Modified = value; _document.MarkPackagePropertiesDirty(); }
        }

        /// <summary>Last printed timestamp (UTC).</summary>
        public DateTime? LastPrinted {
            get { return _spreadsheetDocument.PackageProperties.LastPrinted; }
            set { EnsureCorePropertiesPart(); _spreadsheetDocument.PackageProperties.LastPrinted = value; _document.MarkPackagePropertiesDirty(); }
        }
    }
}
