using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Validation;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Provides test setup for Excel documents.
    /// </summary>
    public partial class Excel {
        private readonly string _directoryDocuments;
        private readonly string _directoryWithFiles;
        private readonly string _directoryWithImages;

        public Excel() {
            _directoryDocuments = Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Documents");
            _directoryWithImages = Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");
            //_directoryDocuments = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "Tests", "TempDocuments");
            _directoryWithFiles = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TempDocuments1");
            Word.Setup(_directoryWithFiles);
        }

        internal static string FormatValidationErrors(IEnumerable<ValidationErrorInfo> errors) {
            return string.Join(Environment.NewLine + Environment.NewLine,
                errors.Select(error =>
                    $"Description: {error.Description}\n" +
                    $"Id: {error.Id}\n" +
                    $"ErrorType: {error.ErrorType}\n" +
                    $"Part: {error.Part?.Uri}\n" +
                    $"Path: {error.Path?.XPath}"));
        }
    }
}