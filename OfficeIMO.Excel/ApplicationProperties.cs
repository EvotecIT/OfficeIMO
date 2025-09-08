using System;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Extended (application) properties for an Excel workbook, e.g. Company, Manager.
    /// </summary>
    public sealed class ApplicationProperties {
        private readonly SpreadsheetDocument _spreadsheetDocument;
        private readonly ExcelDocument _document;

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

        public string Company {
            get => _spreadsheetDocument.ExtendedFilePropertiesPart?.Properties?.Company?.Text ?? string.Empty;
            set { var p = EnsureProperties(); p.Company ??= new Company(); p.Company.Text = value; }
        }

        public string Manager {
            get => _spreadsheetDocument.ExtendedFilePropertiesPart?.Properties?.Manager?.Text ?? string.Empty;
            set { var p = EnsureProperties(); p.Manager ??= new Manager(); p.Manager.Text = value; }
        }

        public string ApplicationName {
            get => _spreadsheetDocument.ExtendedFilePropertiesPart?.Properties?.Application?.Text ?? string.Empty;
            set { var p = EnsureProperties(); p.Application ??= new Application(); p.Application.Text = value; }
        }
    }
}

