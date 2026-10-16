using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Extended (application) properties for an Excel workbook, e.g. Company, Manager.
    /// The wrapper remains associated with its workbook across saves.
    /// </summary>
    public sealed class ExcelApplicationProperties {
        private SpreadsheetDocument Document => _document._spreadSheetDocument
            ?? throw new ObjectDisposedException(nameof(ExcelDocument));
        private readonly ExcelDocument _document;

        /// <summary>
        /// Creates a new wrapper for extended (application) properties bound to the given document.
        /// </summary>
        public ExcelApplicationProperties(ExcelDocument document) {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            document.ApplicationProperties = this;
        }

        private Properties EnsureProperties() {
            var part = Document.ExtendedFilePropertiesPart;
            if (part == null) part = Document.AddExtendedFilePropertiesPart();
            if (part.Properties == null) part.Properties = new Properties();
            return part.Properties;
        }

        /// <summary>Company property stored in the Extended File Properties part.</summary>
        public string Company {
            get => Document.ExtendedFilePropertiesPart?.Properties?.Company?.Text ?? string.Empty;
            set { var p = EnsureProperties(); p.Company ??= new Company(); p.Company.Text = value; _document.MarkPackagePropertiesDirty(); }
        }

        /// <summary>Manager property stored in the Extended File Properties part.</summary>
        public string Manager {
            get => Document.ExtendedFilePropertiesPart?.Properties?.Manager?.Text ?? string.Empty;
            set { var p = EnsureProperties(); p.Manager ??= new Manager(); p.Manager.Text = value; _document.MarkPackagePropertiesDirty(); }
        }

        /// <summary>Application name (producer) stored in the Extended File Properties part.</summary>
        public string ApplicationName {
            get => Document.ExtendedFilePropertiesPart?.Properties?.Application?.Text ?? string.Empty;
            set { var p = EnsureProperties(); p.Application ??= new Application(); p.Application.Text = value; _document.MarkPackagePropertiesDirty(); }
        }
    }
}
