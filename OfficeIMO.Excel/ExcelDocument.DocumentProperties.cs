using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Internal;
using System.IO.Packaging;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System;
using System.Diagnostics;
using System.IO;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument : IDisposable, IAsyncDisposable {

        /// <summary>
        /// Path to the file backing this document.
        /// </summary>
        public string? FilePath { get; internal set; }

        /// <summary>
        /// Built-in (core) document properties (Title, Creator, etc.).
        /// </summary>
        public BuiltinDocumentProperties BuiltinDocumentProperties { get; internal set; } = null!;

        /// <summary>
        /// Extended (application) properties (Company, Manager, etc.).
        /// </summary>
        public ApplicationProperties ApplicationProperties { get; internal set; } = null!;

        /// <summary>
        /// Custom workbook properties keyed by property name.
        /// </summary>
        public ExcelCustomDocumentPropertyCollection CustomDocumentProperties { get; } = new ExcelCustomDocumentPropertyCollection();

        /// <summary>
        /// FileOpenAccess of the document
        /// </summary>
        public FileAccess FileOpenAccess => _spreadSheetDocument.FileOpenAccess;

        /// <summary>Gets whether the workbook was loaded for reading or editing.</summary>
        public DocumentAccessMode AccessMode => FileOpenAccess == FileAccess.Read
            ? DocumentAccessMode.ReadOnly
            : DocumentAccessMode.ReadWrite;

        /// <summary>Gets the persistence policy selected when the workbook was created or loaded.</summary>
        public DocumentPersistenceMode PersistenceMode => _persistenceMode;

        /// <summary>
        /// Indicates whether the document is valid.
        /// </summary>
        public bool DocumentIsValid {
            get {
                if (DocumentValidationErrors.Count > 0) {
                    return false;
                }

                return true;
            }
        }

        /// <summary>
        /// Gets the list of validation errors for the document.
        /// </summary>
        public List<ValidationErrorInfo> DocumentValidationErrors {
            get {
                return ValidateDocument();
            }
        }
    }
}
