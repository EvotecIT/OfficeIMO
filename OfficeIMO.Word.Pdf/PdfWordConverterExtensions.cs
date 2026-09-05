using PdfCore = OfficeIMO.Pdf;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Pdf {
    /// <summary>
    /// Converts the first-party logical PDF model into an editable Word document.
    /// PDF parsing, stream handling, and page selection remain owned by <c>OfficeIMO.Pdf</c>.
    /// </summary>
    public static class PdfWordConverterExtensions {
        /// <summary>Converts an opened PDF into an editable Word document.</summary>
        public static WordDocument ToWordDocument(
            this PdfCore.PdfDocument document,
            PdfToWordOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (document == null) throw new ArgumentNullException(nameof(document));
            return ReadForWord(document, options, cancellationToken).ToWordDocument(options, cancellationToken);
        }

        /// <summary>Converts an opened PDF into an editable Word document with conversion diagnostics.</summary>
        public static PdfWordConversionResult ToWordDocumentResult(
            this PdfCore.PdfDocument document,
            PdfToWordOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (document == null) throw new ArgumentNullException(nameof(document));
            return ReadForWord(document, options, cancellationToken).ToWordDocumentResult(options, cancellationToken);
        }

        /// <summary>Converts an opened PDF and saves the editable Word document to a file.</summary>
        public static OfficeOutputResult<PdfWordConversionReport> SaveAsWord(
            this PdfCore.PdfDocument document,
            string path,
            PdfToWordOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (document == null) throw new ArgumentNullException(nameof(document));
            return ReadForWord(document, options, cancellationToken).SaveAsWord(path, options, cancellationToken);
        }

        /// <summary>Converts an opened PDF and saves the editable Word document to a caller-owned stream.</summary>
        public static OfficeOutputResult<PdfWordConversionReport> SaveAsWord(
            this PdfCore.PdfDocument document,
            Stream stream,
            PdfToWordOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (document == null) throw new ArgumentNullException(nameof(document));
            return ReadForWord(document, options, cancellationToken).SaveAsWord(stream, options, cancellationToken);
        }

        /// <summary>Converts an opened PDF and asynchronously saves the editable Word document to a file.</summary>
        public static async Task<OfficeOutputResult<PdfWordConversionReport>> SaveAsWordAsync(
            this PdfCore.PdfDocument document,
            string path,
            PdfToWordOptions? options = null,
            CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (document == null) throw new ArgumentNullException(nameof(document));
            PdfToWordOptions operation = (options ?? new PdfToWordOptions()).CloneForConversion();
            CancellationToken effectiveCancellationToken = cancellationToken;
            operation.CancellationToken = effectiveCancellationToken;
            return await ReadForWord(document, operation, effectiveCancellationToken)
                .SaveAsWordAsync(path, operation, effectiveCancellationToken)
                .ConfigureAwait(false);
        }

        /// <summary>Converts an opened PDF and asynchronously saves the editable Word document to a caller-owned stream.</summary>
        public static async Task<OfficeOutputResult<PdfWordConversionReport>> SaveAsWordAsync(
            this PdfCore.PdfDocument document,
            Stream stream,
            PdfToWordOptions? options = null,
            CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (document == null) throw new ArgumentNullException(nameof(document));
            PdfToWordOptions operation = (options ?? new PdfToWordOptions()).CloneForConversion();
            CancellationToken effectiveCancellationToken = cancellationToken;
            operation.CancellationToken = effectiveCancellationToken;
            return await ReadForWord(document, operation, effectiveCancellationToken)
                .SaveAsWordAsync(stream, operation, effectiveCancellationToken)
                .ConfigureAwait(false);
        }

        private static PdfCore.PdfDocumentReadResult ReadForWord(
            PdfCore.PdfDocument document,
            PdfToWordOptions? options,
            CancellationToken cancellationToken = default) {
            return document.Read(options?.ReadOptions, cancellationToken);
        }

        /// <summary>Converts a logical PDF model into an editable Word document.</summary>
        public static WordDocument ToWordDocument(
            this PdfCore.PdfDocumentReadResult document,
            PdfToWordOptions? options = null, System.Threading.CancellationToken cancellationToken = default) => document.ToWordDocumentResult(options, cancellationToken).Value;

        /// <summary>Converts a logical PDF model into an editable Word document with conversion diagnostics.</summary>
        public static PdfWordConversionResult ToWordDocumentResult(
            this PdfCore.PdfDocumentReadResult document,
            PdfToWordOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (document == null) throw new ArgumentNullException(nameof(document));

            PdfToWordOptions operation = (options ?? new PdfToWordOptions()).CloneForConversion();
        operation.CancellationToken = cancellationToken;
            WordDocument word = PdfWordConverter.Convert(document, operation);
            return new PdfWordConversionResult(word, operation.Report);
        }

        /// <summary>Converts a logical PDF model and saves the editable Word document to a file.</summary>
        public static OfficeOutputResult<PdfWordConversionReport> SaveAsWord(
            this PdfCore.PdfDocumentReadResult document,
            string path,
            PdfToWordOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Document path cannot be empty.", nameof(path));
            PdfWordConversionResult result = document.ToWordDocumentResult(options, cancellationToken);
            using (result.Value) {
                cancellationToken.ThrowIfCancellationRequested();
                result.Value.Save(path);
            }
            return OfficeOutputResult<PdfWordConversionReport>.FromSuccess(path, result.Report);
        }

        /// <summary>Converts a logical PDF model and saves the editable Word document to a caller-owned stream.</summary>
        public static OfficeOutputResult<PdfWordConversionReport> SaveAsWord(
            this PdfCore.PdfDocumentReadResult document,
            Stream stream,
            PdfToWordOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(stream));
            PdfWordConversionResult result = document.ToWordDocumentResult(options, cancellationToken);
            using (result.Value) {
                cancellationToken.ThrowIfCancellationRequested();
                result.Value.Save(stream);
            }
            return OfficeOutputResult<PdfWordConversionReport>.FromSuccess(null, result.Report);
        }

        /// <summary>Converts a logical PDF model and asynchronously saves the editable Word document to a file.</summary>
        public static async Task<OfficeOutputResult<PdfWordConversionReport>> SaveAsWordAsync(
            this PdfCore.PdfDocumentReadResult document,
            string path,
            PdfToWordOptions? options = null,
            CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Document path cannot be empty.", nameof(path));
            PdfToWordOptions operation = (options ?? new PdfToWordOptions()).CloneForConversion();
            CancellationToken effectiveCancellationToken = cancellationToken;
            operation.CancellationToken = effectiveCancellationToken;
            operation.CancellationToken.ThrowIfCancellationRequested();
            PdfWordConversionResult result = document.ToWordDocumentResult(operation, cancellationToken);
            using (result.Value) {
                await result.Value.SaveAsync(path, operation.CancellationToken).ConfigureAwait(false);
            }
            return OfficeOutputResult<PdfWordConversionReport>.FromSuccess(path, result.Report);
        }

        /// <summary>Converts a logical PDF model and asynchronously saves the editable Word document to a caller-owned stream.</summary>
        public static async Task<OfficeOutputResult<PdfWordConversionReport>> SaveAsWordAsync(
            this PdfCore.PdfDocumentReadResult document,
            Stream stream,
            PdfToWordOptions? options = null,
            CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(stream));
            PdfToWordOptions operation = (options ?? new PdfToWordOptions()).CloneForConversion();
            CancellationToken effectiveCancellationToken = cancellationToken;
            operation.CancellationToken = effectiveCancellationToken;
            operation.CancellationToken.ThrowIfCancellationRequested();
            PdfWordConversionResult result = document.ToWordDocumentResult(operation, cancellationToken);
            using (result.Value) {
                await result.Value.SaveAsync(stream, operation.CancellationToken).ConfigureAwait(false);
            }
            return OfficeOutputResult<PdfWordConversionReport>.FromSuccess(null, result.Report);
        }


    }
}
