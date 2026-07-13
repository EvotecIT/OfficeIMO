using PdfCore = OfficeIMO.Pdf;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Pdf {
    /// <summary>
    /// Converts the first-party logical PDF model into an editable Word document.
    /// PDF parsing, stream handling, and page selection remain owned by <c>OfficeIMO.Pdf</c>.
    /// </summary>
    public static class PdfWordConverterExtensions {
        /// <summary>Converts a logical PDF model into an editable Word document.</summary>
        public static WordDocument ToWordDocument(
            this PdfCore.PdfLogicalDocument document,
            PdfWordReadOptions? options = null) => document.ToWordDocumentResult(options).Value;

        /// <summary>Converts a logical PDF model into an editable Word document with conversion diagnostics.</summary>
        public static PdfWordConversionResult ToWordDocumentResult(
            this PdfCore.PdfLogicalDocument document,
            PdfWordReadOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));

            PdfWordReadOptions operation = (options ?? new PdfWordReadOptions()).CloneForConversion();
            WordDocument word = PdfWordConverter.Convert(document, operation);
            return new PdfWordConversionResult(word, operation.Report);
        }

        /// <summary>Converts a logical PDF model and saves the editable Word document to a file.</summary>
        public static PdfWordConversionReport SaveAsWord(
            this PdfCore.PdfLogicalDocument document,
            string path,
            PdfWordReadOptions? options = null) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Document path cannot be empty.", nameof(path));
            PdfWordConversionResult result = document.ToWordDocumentResult(options);
            using (result.Value) {
                result.Value.Save(path);
            }
            return result.Report;
        }

        /// <summary>Converts a logical PDF model and saves the editable Word document to a caller-owned stream.</summary>
        public static PdfWordConversionReport SaveAsWord(
            this PdfCore.PdfLogicalDocument document,
            Stream stream,
            PdfWordReadOptions? options = null) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(stream));
            PdfWordConversionResult result = document.ToWordDocumentResult(options);
            using (result.Value) {
                result.Value.Save(stream);
            }
            return result.Report;
        }

        /// <summary>Converts a logical PDF model and asynchronously saves the editable Word document to a file.</summary>
        public static async Task<PdfWordConversionReport> SaveAsWordAsync(
            this PdfCore.PdfLogicalDocument document,
            string path,
            PdfWordReadOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Document path cannot be empty.", nameof(path));
            cancellationToken.ThrowIfCancellationRequested();
            PdfWordConversionResult result = document.ToWordDocumentResult(options);
            using (result.Value) {
                await result.Value.SaveAsync(path, cancellationToken).ConfigureAwait(false);
            }
            return result.Report;
        }

        /// <summary>Converts a logical PDF model and asynchronously saves the editable Word document to a caller-owned stream.</summary>
        public static async Task<PdfWordConversionReport> SaveAsWordAsync(
            this PdfCore.PdfLogicalDocument document,
            Stream stream,
            PdfWordReadOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(stream));
            cancellationToken.ThrowIfCancellationRequested();
            PdfWordConversionResult result = document.ToWordDocumentResult(options);
            using (result.Value) {
                await result.Value.SaveAsync(stream, cancellationToken).ConfigureAwait(false);
            }
            return result.Report;
        }
    }
}
