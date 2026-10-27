using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {

    /// <summary>
    /// Provides extension methods for converting <see cref="WordDocument"/> instances to PDF files.
    /// </summary>
    public static partial class WordPdfConverterExtensions {


        /// <summary>
        /// Converts the specified <see cref="WordDocument"/> to a PDF document and returns conversion diagnostics with it.
        /// </summary>
        internal static PdfCore.PdfDocumentConversionResult ConvertToPdfWithEvidence(this WordDocument document, WordPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            WordPdfExportOptions operation = (options ?? new WordPdfExportOptions()).CloneForConversion();
            operation.CancellationToken = cancellationToken;
            operation.CancellationToken.ThrowIfCancellationRequested();
            if (operation.DateTimeFieldUpdateOptions != null) {
                document.UpdateFieldsAndGetReport(operation.DateTimeFieldUpdateOptions);
                operation.CancellationToken.ThrowIfCancellationRequested();
            }
            PdfCore.PdfDocument pdf;
            using IDisposable runPaintScope = WordRunFormatting.BeginConversionScope(cancellationToken);
            using IDisposable storyScope = WordHeaderFooterResolutionSnapshot.Begin(document, cancellationToken);
            using (WordComplexFieldRunVisibility.BeginConversionScope())
                pdf = CreateOfficeIMOPdfDocument(document, operation);
            operation.CancellationToken.ThrowIfCancellationRequested();
            return new PdfCore.PdfDocumentConversionResult(pdf, operation.Report);
        }





















        private static string ValidateOutputPath(string path, string paramName) {
            string fullPath;
            try {
                fullPath = Path.GetFullPath(path);
            } catch (Exception ex) {
                throw new ArgumentException("Path is invalid.", paramName, ex);
            }

            if (Directory.Exists(fullPath) && (File.GetAttributes(fullPath) & FileAttributes.Directory) == FileAttributes.Directory) {
                throw new ArgumentException("Path refers to a directory; a file path is required.", paramName);
            }

            string fileName = Path.GetFileName(fullPath);
            if (string.IsNullOrEmpty(fileName)) {
                throw new ArgumentException("Path must include a file name.", paramName);
            }

            if (fileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) {
                throw new ArgumentException("Path contains invalid file name characters.", paramName);
            }

            return fullPath;
        }
    }
}
