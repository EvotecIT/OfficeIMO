using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.Core.Internal;
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
        /// Validates the document using the specified file format version.
        /// </summary>
        /// <param name="fileFormatVersion">File format version to validate against.</param>
        /// <returns>List of validation errors.</returns>
        public List<OfficeOpenXmlValidationError> ValidateDocument(
            OfficeOpenXmlFileFormatVersion fileFormatVersion = OfficeOpenXmlFileFormatVersion.Microsoft365) {
            var listErrors = new List<OfficeOpenXmlValidationError>();
            OpenXmlValidator validator = new OpenXmlValidator(fileFormatVersion.ToOpenXml());
            foreach (ValidationErrorInfo error in validator.Validate(_spreadSheetDocument)) {
                listErrors.Add(error.ToOfficeValidationError());
            }
            return listErrors;
        }
    }
}
