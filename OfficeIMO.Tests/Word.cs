using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Office.CustomUI;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    [Collection("WordTests")]
    public partial class Word {
        private readonly string _directoryDocuments;
        private readonly string _directoryWithFiles;
        private readonly string _directoryWithImages;

        internal static void Setup(string path) {
            if (!Directory.Exists(path)) {
                Directory.CreateDirectory(path);
            } else {
                //Directory.Delete(path, true);
                //Directory.CreateDirectory(path);
            }
        }

        public Word() {
            _directoryDocuments = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents");
            _directoryWithImages = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images");
            // Create a unique per-test directory to avoid parallel write collisions
            string unique = Guid.NewGuid().ToString("N");
            _directoryWithFiles = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TempDocuments2", unique);
            Setup(_directoryWithFiles);
        }

        /// <summary>
        /// Copies a fixture document from the read-only Documents folder into the current
        /// test's unique working directory and returns the destination path. Safe for
        /// parallel execution and ideal for scenarios where the test saves modifications.
        /// </summary>
        /// <param name="fileName">Fixture file name that exists under the Documents folder.</param>
        /// <returns>Absolute path to the copied file in the test's TempDocuments2 folder.</returns>
        protected string CopyFixtureDoc(string fileName) {
            string source = Path.Combine(_directoryDocuments, fileName);
            string dest = Path.Combine(_directoryWithFiles, fileName);
            Directory.CreateDirectory(Path.GetDirectoryName(dest)!);
            File.Copy(source, dest, overwrite: true);
            return dest;
        }

        /// <summary>
        /// Returns the absolute path to a fixture under Documents. Use only for read-only
        /// access; prefer <see cref="CopyFixtureDoc"/> if the test will modify or save.
        /// </summary>
        protected string GetFixtureDoc(string fileName) => Path.Combine(_directoryDocuments, fileName);

        /// <summary>
        /// This helps finding unexpected elements during validation. Should prevent unexpected changes
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public bool HasUnexpectedElements(WordDocument document) {
            bool found = false;
            foreach (var e in document.DocumentValidationErrors) {
                if (e.Description.StartsWith("The element has unexpected child element")) {
                    found = true;
                    break;
                }
            }
            return found;
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
