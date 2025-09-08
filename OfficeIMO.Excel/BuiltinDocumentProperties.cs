using System;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Built-in (core) document properties for an Excel workbook, such as Title and Author.
    /// </summary>
    public sealed class BuiltinDocumentProperties {
        private readonly SpreadsheetDocument _spreadsheetDocument;
        private readonly ExcelDocument _document;

        public BuiltinDocumentProperties(ExcelDocument document) {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _spreadsheetDocument = document._spreadSheetDocument;
            document.BuiltinDocumentProperties = this;
            EnsureCorePropertiesPart();
        }

        private void EnsureCorePropertiesPart()
        {
            // Touch the package properties to ensure the backing object is initialized.
            try { var _ = _spreadsheetDocument.PackageProperties; } catch { }
        }

        public string? Creator {
            get { return _spreadsheetDocument.PackageProperties.Creator; }
            set { EnsureCorePropertiesPart(); _spreadsheetDocument.PackageProperties.Creator = value; }
        }

        public string? Title {
            get { return _spreadsheetDocument.PackageProperties.Title; }
            set { EnsureCorePropertiesPart(); _spreadsheetDocument.PackageProperties.Title = value; }
        }

        public string? Description {
            get { return _spreadsheetDocument.PackageProperties.Description; }
            set { EnsureCorePropertiesPart(); _spreadsheetDocument.PackageProperties.Description = value; }
        }

        public string? Category {
            get { return _spreadsheetDocument.PackageProperties.Category; }
            set { EnsureCorePropertiesPart(); _spreadsheetDocument.PackageProperties.Category = value; }
        }

        public string? Keywords {
            get { return _spreadsheetDocument.PackageProperties.Keywords; }
            set { EnsureCorePropertiesPart(); _spreadsheetDocument.PackageProperties.Keywords = value; }
        }

        public string? Subject {
            get { return _spreadsheetDocument.PackageProperties.Subject; }
            set { EnsureCorePropertiesPart(); _spreadsheetDocument.PackageProperties.Subject = value; }
        }

        public string? Revision {
            get { return _spreadsheetDocument.PackageProperties.Revision; }
            set { EnsureCorePropertiesPart(); _spreadsheetDocument.PackageProperties.Revision = value; }
        }

        public string? LastModifiedBy {
            get { return _spreadsheetDocument.PackageProperties.LastModifiedBy; }
            set { EnsureCorePropertiesPart(); _spreadsheetDocument.PackageProperties.LastModifiedBy = value; }
        }

        public string? Version {
            get { return _spreadsheetDocument.PackageProperties.Version; }
            set { EnsureCorePropertiesPart(); _spreadsheetDocument.PackageProperties.Version = value; }
        }

        public DateTime? Created {
            get { return _spreadsheetDocument.PackageProperties.Created; }
            set { EnsureCorePropertiesPart(); _spreadsheetDocument.PackageProperties.Created = value; }
        }

        public DateTime? Modified {
            get { return _spreadsheetDocument.PackageProperties.Modified; }
            set { EnsureCorePropertiesPart(); _spreadsheetDocument.PackageProperties.Modified = value; }
        }

        public DateTime? LastPrinted {
            get { return _spreadsheetDocument.PackageProperties.LastPrinted; }
            set { EnsureCorePropertiesPart(); _spreadsheetDocument.PackageProperties.LastPrinted = value; }
        }
    }
}
