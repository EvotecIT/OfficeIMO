using PdfCore = OfficeIMO.Pdf;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Pdf.Word {
    /// <summary>
    /// Converts PDF documents and logical content into Word documents.
    /// PDF parsing, stream handling, and page selection remain owned by <c>OfficeIMO.Pdf</c>.
    /// </summary>
    public static partial class PdfWordConverterExtensions {


        /// <summary>Converts an opened PDF into a Word document using the selected import mode with conversion diagnostics.</summary>
        internal static PdfWordConversionResult ConvertToWordWithEvidence(
            this PdfCore.PdfDocument document,
            WordPdfImportOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (document == null) throw new ArgumentNullException(nameof(document));
            WordPdfImportOptions operation = (options ?? new WordPdfImportOptions()).CloneForConversion();
            operation.CancellationToken = cancellationToken;
            ValidateImportMode(operation.Mode);
            return operation.Mode == WordPdfImportMode.VisualPages
                ? PdfWordConverter.ConvertVisualPages(document, operation)
                : ReadForWord(document, operation, cancellationToken).ConvertToWordWithEvidence(operation, cancellationToken);
        }









        private static PdfCore.PdfDocumentReadResult ReadForWord(
            PdfCore.PdfDocument document,
            WordPdfImportOptions? options,
            CancellationToken cancellationToken = default) {
            return document.Read(options?.ReadOptions, cancellationToken);
        }



        /// <summary>Converts a logical PDF model into an editable Word document with conversion diagnostics.</summary>
        internal static PdfWordConversionResult ConvertToWordWithEvidence(
            this PdfCore.PdfDocumentReadResult document,
            WordPdfImportOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (document == null) throw new ArgumentNullException(nameof(document));

            WordPdfImportOptions operation = (options ?? new WordPdfImportOptions()).CloneForConversion();
            ValidateImportMode(operation.Mode);
            if (operation.Mode != WordPdfImportMode.EditableContent)
                throw new NotSupportedException("Visual page conversion requires an opened PDF document, not a reduced logical model.");
        operation.CancellationToken = cancellationToken;
            WordDocument word = PdfWordConverter.Convert(document, operation);
            return new PdfWordConversionResult(word, operation.Report);
        }








        private static void ValidateImportMode(WordPdfImportMode mode) {
            if (mode < WordPdfImportMode.EditableContent || mode > WordPdfImportMode.VisualPages)
                throw new ArgumentOutOfRangeException(nameof(mode));
        }
    }
}
