using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Extended (application) properties for an Excel workbook, e.g. Company, Manager.
    /// </summary>
    public sealed class ApplicationProperties {
        private readonly SpreadsheetDocument _spreadsheetDocument;
        private readonly ExcelDocument _document;

        /// <summary>
        /// Creates a new wrapper for extended (application) properties bound to the given document.
        /// </summary>
        public ApplicationProperties(ExcelDocument document) {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _spreadsheetDocument = document._spreadSheetDocument;
            document.ApplicationProperties = this;
        }

        private Properties EnsureProperties() {
            var part = _spreadsheetDocument.ExtendedFilePropertiesPart;
            if (part == null) part = _spreadsheetDocument.AddExtendedFilePropertiesPart();
            if (part.Properties == null) part.Properties = new Properties();
            return part.Properties;
        }

        /// <summary>Company property stored in the Extended File Properties part.</summary>
        public string Company {
            get => _spreadsheetDocument.ExtendedFilePropertiesPart?.Properties?.Company?.Text ?? string.Empty;
            set { var p = EnsureProperties(); p.Company ??= new Company(); p.Company.Text = value; }
        }

        /// <summary>Manager property stored in the Extended File Properties part.</summary>
        public string Manager {
            get => _spreadsheetDocument.ExtendedFilePropertiesPart?.Properties?.Manager?.Text ?? string.Empty;
            set { var p = EnsureProperties(); p.Manager ??= new Manager(); p.Manager.Text = value; }
        }

        /// <summary>Application name (producer) stored in the Extended File Properties part.</summary>
        public string ApplicationName {
            get => _spreadsheetDocument.ExtendedFilePropertiesPart?.Properties?.Application?.Text ?? string.Empty;
            set { var p = EnsureProperties(); p.Application ??= new Application(); p.Application.Text = value; }
        }
    }
}
