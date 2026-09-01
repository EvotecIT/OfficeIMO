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
            PdfWordImportOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return ReadForWord(document, options).ToWordDocument(options);
        }

        /// <summary>Converts an opened PDF into an editable Word document with conversion diagnostics.</summary>
        public static PdfWordConversionResult ToWordDocumentResult(
            this PdfCore.PdfDocument document,
            PdfWordImportOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return ReadForWord(document, options).ToWordDocumentResult(options);
        }

        /// <summary>Converts an opened PDF and saves the editable Word document to a file.</summary>
        public static PdfWordConversionReport SaveAsWord(
            this PdfCore.PdfDocument document,
            string path,
            PdfWordImportOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return ReadForWord(document, options).SaveAsWord(path, options);
        }

        /// <summary>Converts an opened PDF and saves the editable Word document to a caller-owned stream.</summary>
        public static PdfWordConversionReport SaveAsWord(
            this PdfCore.PdfDocument document,
            Stream stream,
            PdfWordImportOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return ReadForWord(document, options).SaveAsWord(stream, options);
        }

        /// <summary>Converts an opened PDF and asynchronously saves the editable Word document to a file.</summary>
        public static async Task<PdfWordConversionReport> SaveAsWordAsync(
            this PdfCore.PdfDocument document,
            string path,
            PdfWordImportOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            PdfWordImportOptions operation = (options ?? new PdfWordImportOptions()).CloneForConversion();
            using CancellationTokenSource? linked = LinkCancellationTokens(operation.CancellationToken, cancellationToken, out CancellationToken effectiveCancellationToken);
            operation.CancellationToken = effectiveCancellationToken;
            return await ReadForWord(document, operation, effectiveCancellationToken)
                .SaveAsWordAsync(path, operation)
                .ConfigureAwait(false);
        }

        /// <summary>Converts an opened PDF and asynchronously saves the editable Word document to a caller-owned stream.</summary>
        public static async Task<PdfWordConversionReport> SaveAsWordAsync(
            this PdfCore.PdfDocument document,
            Stream stream,
            PdfWordImportOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            PdfWordImportOptions operation = (options ?? new PdfWordImportOptions()).CloneForConversion();
            using CancellationTokenSource? linked = LinkCancellationTokens(operation.CancellationToken, cancellationToken, out CancellationToken effectiveCancellationToken);
            operation.CancellationToken = effectiveCancellationToken;
            return await ReadForWord(document, operation, effectiveCancellationToken)
                .SaveAsWordAsync(stream, operation)
                .ConfigureAwait(false);
        }

        private static PdfCore.PdfDocumentReadResult ReadForWord(
            PdfCore.PdfDocument document,
            PdfWordImportOptions? options,
            CancellationToken cancellationToken = default) {
            CancellationToken effectiveCancellationToken = cancellationToken.CanBeCanceled
                ? cancellationToken
                : options?.CancellationToken ?? default;
            return document.Read(options?.ReadOptions, effectiveCancellationToken);
        }

        /// <summary>Converts a logical PDF model into an editable Word document.</summary>
        public static WordDocument ToWordDocument(
            this PdfCore.PdfDocumentReadResult document,
            PdfWordImportOptions? options = null) => document.ToWordDocumentResult(options).Value;

        /// <summary>Converts a logical PDF model into an editable Word document with conversion diagnostics.</summary>
        public static PdfWordConversionResult ToWordDocumentResult(
            this PdfCore.PdfDocumentReadResult document,
            PdfWordImportOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));

            PdfWordImportOptions operation = (options ?? new PdfWordImportOptions()).CloneForConversion();
            WordDocument word = PdfWordConverter.Convert(document, operation);
            return new PdfWordConversionResult(word, operation.Report);
        }

        /// <summary>Converts a logical PDF model and saves the editable Word document to a file.</summary>
        public static PdfWordConversionReport SaveAsWord(
            this PdfCore.PdfDocumentReadResult document,
            string path,
            PdfWordImportOptions? options = null) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Document path cannot be empty.", nameof(path));
            PdfWordConversionResult result = document.ToWordDocumentResult(options);
            using (result.Value) {
                result.Value.Save(path);
            }
            return result.Report;
        }

        /// <summary>Converts a logical PDF model and saves the editable Word document to a caller-owned stream.</summary>
        public static PdfWordConversionReport SaveAsWord(
            this PdfCore.PdfDocumentReadResult document,
            Stream stream,
            PdfWordImportOptions? options = null) {
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
            this PdfCore.PdfDocumentReadResult document,
            string path,
            PdfWordImportOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Document path cannot be empty.", nameof(path));
            PdfWordImportOptions operation = (options ?? new PdfWordImportOptions()).CloneForConversion();
            using CancellationTokenSource? linked = LinkCancellationTokens(operation.CancellationToken, cancellationToken, out CancellationToken effectiveCancellationToken);
            operation.CancellationToken = effectiveCancellationToken;
            operation.CancellationToken.ThrowIfCancellationRequested();
            PdfWordConversionResult result = document.ToWordDocumentResult(operation);
            using (result.Value) {
                await result.Value.SaveAsync(path, operation.CancellationToken).ConfigureAwait(false);
            }
            return result.Report;
        }

        /// <summary>Converts a logical PDF model and asynchronously saves the editable Word document to a caller-owned stream.</summary>
        public static async Task<PdfWordConversionReport> SaveAsWordAsync(
            this PdfCore.PdfDocumentReadResult document,
            Stream stream,
            PdfWordImportOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(stream));
            PdfWordImportOptions operation = (options ?? new PdfWordImportOptions()).CloneForConversion();
            using CancellationTokenSource? linked = LinkCancellationTokens(operation.CancellationToken, cancellationToken, out CancellationToken effectiveCancellationToken);
            operation.CancellationToken = effectiveCancellationToken;
            operation.CancellationToken.ThrowIfCancellationRequested();
            PdfWordConversionResult result = document.ToWordDocumentResult(operation);
            using (result.Value) {
                await result.Value.SaveAsync(stream, operation.CancellationToken).ConfigureAwait(false);
            }
            return result.Report;
        }

        private static CancellationTokenSource? LinkCancellationTokens(
            CancellationToken optionsToken,
            CancellationToken methodToken,
            out CancellationToken effectiveToken) {
            if (optionsToken.CanBeCanceled && methodToken.CanBeCanceled && optionsToken != methodToken) {
                CancellationTokenSource linked = CancellationTokenSource.CreateLinkedTokenSource(optionsToken, methodToken);
                effectiveToken = linked.Token;
                return linked;
            }
            effectiveToken = methodToken.CanBeCanceled ? methodToken : optionsToken;
            return null;
        }
    }
}
