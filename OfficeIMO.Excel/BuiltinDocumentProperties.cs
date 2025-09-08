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
        }

        public string? Creator {
            get => _spreadsheetDocument.PackageProperties.Creator;
            set => _spreadsheetDocument.PackageProperties.Creator = value;
        }

        public string? Title {
            get => _spreadsheetDocument.PackageProperties.Title;
            set => _spreadsheetDocument.PackageProperties.Title = value;
        }

        public string? Description {
            get => _spreadsheetDocument.PackageProperties.Description;
            set => _spreadsheetDocument.PackageProperties.Description = value;
        }

        public string? Category {
            get => _spreadsheetDocument.PackageProperties.Category;
            set => _spreadsheetDocument.PackageProperties.Category = value;
        }

        public string? Keywords {
            get => _spreadsheetDocument.PackageProperties.Keywords;
            set => _spreadsheetDocument.PackageProperties.Keywords = value;
        }

        public string? Subject {
            get => _spreadsheetDocument.PackageProperties.Subject;
            set => _spreadsheetDocument.PackageProperties.Subject = value;
        }

        public string? Revision {
            get => _spreadsheetDocument.PackageProperties.Revision;
            set => _spreadsheetDocument.PackageProperties.Revision = value;
        }

        public string? LastModifiedBy {
            get => _spreadsheetDocument.PackageProperties.LastModifiedBy;
            set => _spreadsheetDocument.PackageProperties.LastModifiedBy = value;
        }

        public string? Version {
            get => _spreadsheetDocument.PackageProperties.Version;
            set => _spreadsheetDocument.PackageProperties.Version = value;
        }

        public DateTime? Created {
            get => _spreadsheetDocument.PackageProperties.Created;
            set => _spreadsheetDocument.PackageProperties.Created = value;
        }

        public DateTime? Modified {
            get => _spreadsheetDocument.PackageProperties.Modified;
            set => _spreadsheetDocument.PackageProperties.Modified = value;
        }

        public DateTime? LastPrinted {
            get => _spreadsheetDocument.PackageProperties.LastPrinted;
            set => _spreadsheetDocument.PackageProperties.LastPrinted = value;
        }
    }
}

