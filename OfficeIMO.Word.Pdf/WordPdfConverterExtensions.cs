using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {

    /// <summary>
    /// Provides extension methods for converting <see cref="WordDocument"/> instances to PDF files.
    /// </summary>
    public static partial class WordPdfConverterExtensions {
        /// <summary>
        /// Converts the specified <see cref="WordDocument"/> to a first-party OfficeIMO PDF document model.
        /// </summary>
        /// <param name="document">The document to convert.</param>
        /// <param name="options">Optional PDF configuration.</param>
        /// <param name="cancellationToken">Cancellation observed at document-section and element boundaries.</param>
        /// <returns>The generated first-party PDF document model.</returns>
        public static PdfCore.PdfDocument ToPdfDocument(this WordDocument document, WordToPdfOptions? options = null, CancellationToken cancellationToken = default) {
            return document.ToPdfDocumentResult(options, cancellationToken).Value;
        }

        /// <summary>
        /// Converts the specified <see cref="WordDocument"/> to a PDF document and returns conversion diagnostics with it.
        /// </summary>
        public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this WordDocument document, WordToPdfOptions? options = null, CancellationToken cancellationToken = default) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            WordToPdfOptions operation = (options ?? new WordToPdfOptions()).CloneForConversion();
            operation.CancellationToken = cancellationToken;
            operation.CancellationToken.ThrowIfCancellationRequested();
            PdfCore.PdfDocument pdf = CreateOfficeIMOPdfDocument(document, operation);
            operation.CancellationToken.ThrowIfCancellationRequested();
            return new PdfCore.PdfDocumentConversionResult(pdf, operation.Report);
        }

        /// <summary>
        /// Saves the specified <see cref="WordDocument"/> as a PDF at the given <paramref name="path"/>.
        /// </summary>
        /// <param name="document">The document to convert.</param>
        /// <param name="path">The output PDF file path.</param>
        /// <param name="options">Optional PDF configuration.</param>
        /// <param name="cancellationToken">Cancellation observed during conversion.</param>
        public static PdfCore.PdfSaveResult SaveAsPdf(this WordDocument document, string path, WordToPdfOptions? options = null, CancellationToken cancellationToken = default) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            if (path == null) {
                throw new ArgumentNullException(nameof(path));
            }

            if (string.IsNullOrWhiteSpace(path)) {
                throw new ArgumentException("Path cannot be empty or whitespace.", nameof(path));
            }

            string fullPath = ValidateOutputPath(path, nameof(path));
            string? directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(directory)) {
                Directory.CreateDirectory(directory);
            }

            return document.ToPdfDocumentResult(options, cancellationToken).Save(fullPath);
        }

        /// <summary>
        /// Attempts to save the specified <see cref="WordDocument"/> as a PDF file and returns output diagnostics instead of throwing.
        /// </summary>
        public static PdfCore.PdfSaveResult SaveAsPdfResult(this WordDocument document, string path, WordToPdfOptions? options = null, CancellationToken cancellationToken = default) {
            try {
                if (document == null) {
                    throw new ArgumentNullException(nameof(document));
                }

                return document.ToPdfDocumentResult(options, cancellationToken).SaveResult(path);
            } catch (OperationCanceledException) {
                throw;
            } catch (Exception ex) {
                return PdfCore.PdfSaveResult.FromFailure(path, ex);
            }
        }

        /// <summary>
        /// Saves the specified <see cref="WordDocument"/> as a PDF to the provided <paramref name="stream"/>.
        /// </summary>
        /// <param name="document">The document to convert.</param>
        /// <param name="stream">The output stream to receive the PDF data.</param>
        /// <param name="options">Optional PDF configuration.</param>
        /// <param name="cancellationToken">Cancellation observed during conversion.</param>
        public static PdfCore.PdfSaveResult SaveAsPdf(this WordDocument document, Stream stream, WordToPdfOptions? options = null, CancellationToken cancellationToken = default) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            if (stream == null) {
                throw new ArgumentNullException(nameof(stream));
            }

            if (!stream.CanWrite) {
                throw new ArgumentException("Stream must be writable.", nameof(stream));
            }

            PdfCore.PdfSaveResult result = document.ToPdfDocumentResult(options, cancellationToken).Save(stream);
            if (stream.CanSeek) {
                stream.Position = 0;
            }
            return result;
        }

        /// <summary>
        /// Attempts to write the specified <see cref="WordDocument"/> as a PDF to a stream and returns output diagnostics instead of throwing.
        /// </summary>
        public static PdfCore.PdfSaveResult SaveAsPdfResult(this WordDocument document, Stream stream, WordToPdfOptions? options = null, CancellationToken cancellationToken = default) {
            try {
                if (document == null) {
                    throw new ArgumentNullException(nameof(document));
                }

                PdfCore.PdfSaveResult result = document.ToPdfDocumentResult(options, cancellationToken).SaveResult(stream);
                if (result.Succeeded && stream != null && stream.CanSeek) {
                    stream.Position = 0;
                }

                return result;
            } catch (OperationCanceledException) {
                throw;
            } catch (Exception ex) {
                return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex);
            }
        }

        /// <summary>
        /// Converts the specified <see cref="WordDocument"/> to PDF bytes.
        /// </summary>
        /// <param name="document">The document to convert.</param>
        /// <param name="options">Optional PDF configuration.</param>
        /// <param name="cancellationToken">Cancellation observed during conversion.</param>
        /// <returns>The generated PDF as a byte array.</returns>
        /// <example><code>byte[] pdf = document.ToPdfBytes();</code></example>
        public static byte[] ToPdfBytes(this WordDocument document, WordToPdfOptions? options = null, CancellationToken cancellationToken = default) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            return document.ToPdfDocument(options, cancellationToken).ToBytes();
        }

        /// <summary>
        /// Saves the specified <see cref="WordDocument"/> as a PDF at the given <paramref name="path"/> asynchronously.
        /// </summary>
        /// <param name="document">The document to convert.</param>
        /// <param name="path">The output PDF file path.</param>
        /// <param name="options">Optional PDF configuration.</param>
        /// <param name="cancellationToken">A token to observe while waiting for the task to complete.</param>
        /// <returns>The saved PDF output result with conversion and pipeline diagnostics.</returns>
        public static async Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(this WordDocument document, string path, WordToPdfOptions? options = null, CancellationToken cancellationToken = default) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            if (path == null) {
                throw new ArgumentNullException(nameof(path));
            }

            if (string.IsNullOrWhiteSpace(path)) {
                throw new ArgumentException("Path cannot be empty or whitespace.", nameof(path));
            }

            string fullPath = ValidateOutputPath(path, nameof(path));
            string? directory = Path.GetDirectoryName(fullPath);
            cancellationToken.ThrowIfCancellationRequested();
            if (!string.IsNullOrEmpty(directory)) {
                Directory.CreateDirectory(directory);
            }

            using CancellationTokenSource? linked = CreateAsyncConversionOptions(options, cancellationToken, out WordToPdfOptions operation);
            return await document.ToPdfDocumentResult(operation, operation.CancellationToken).SaveAsync(fullPath, operation.CancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Attempts to save the specified <see cref="WordDocument"/> as a PDF file asynchronously and returns output diagnostics instead of throwing.
        /// </summary>
        public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(this WordDocument document, string path, WordToPdfOptions? options = null, CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            try {
                if (document == null) {
                    throw new ArgumentNullException(nameof(document));
                }

                using CancellationTokenSource? linked = CreateAsyncConversionOptions(options, cancellationToken, out WordToPdfOptions operation);
                return await document.ToPdfDocumentResult(operation, operation.CancellationToken).SaveResultAsync(path, operation.CancellationToken).ConfigureAwait(false);
            } catch (OperationCanceledException) {
                throw;
            } catch (Exception ex) {
                return PdfCore.PdfSaveResult.FromFailure(path, ex);
            }
        }

        /// <summary>
        /// Saves the specified <see cref="WordDocument"/> as a PDF to the provided <paramref name="stream"/> asynchronously.
        /// </summary>
        /// <param name="document">The document to convert.</param>
        /// <param name="stream">The output stream to receive the PDF data.</param>
        /// <param name="options">Optional PDF configuration.</param>
        /// <param name="cancellationToken">A token to observe while waiting for the task to complete.</param>
        /// <returns>The saved PDF output result with conversion and pipeline diagnostics.</returns>
        public static async Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(this WordDocument document, Stream stream, WordToPdfOptions? options = null, CancellationToken cancellationToken = default) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            if (stream == null) {
                throw new ArgumentNullException(nameof(stream));
            }

            cancellationToken.ThrowIfCancellationRequested();

            if (!stream.CanWrite) {
                throw new ArgumentException("Stream must be writable.", nameof(stream));
            }

            using CancellationTokenSource? linked = CreateAsyncConversionOptions(options, cancellationToken, out WordToPdfOptions operation);
            PdfCore.PdfSaveResult result = await document.ToPdfDocumentResult(operation, operation.CancellationToken).SaveAsync(stream, operation.CancellationToken).ConfigureAwait(false);
            if (stream.CanSeek) {
                stream.Position = 0;
            }
            return result;
        }

        /// <summary>
        /// Attempts to write the specified <see cref="WordDocument"/> as a PDF to a stream asynchronously and returns output diagnostics instead of throwing.
        /// </summary>
        public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(this WordDocument document, Stream stream, WordToPdfOptions? options = null, CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            try {
                if (document == null) {
                    throw new ArgumentNullException(nameof(document));
                }

                using CancellationTokenSource? linked = CreateAsyncConversionOptions(options, cancellationToken, out WordToPdfOptions operation);
                PdfCore.PdfSaveResult result = await document.ToPdfDocumentResult(operation, operation.CancellationToken).SaveResultAsync(stream, operation.CancellationToken).ConfigureAwait(false);
                if (result.Succeeded && stream != null && stream.CanSeek) {
                    stream.Position = 0;
                }

                return result;
            } catch (OperationCanceledException) {
                throw;
            } catch (Exception ex) {
                return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex);
            }
        }

        private static CancellationTokenSource? CreateAsyncConversionOptions(
            WordToPdfOptions? options,
            CancellationToken methodToken,
            out WordToPdfOptions operation) {
            operation = (options ?? new WordToPdfOptions()).CloneForConversion();
            if (!methodToken.CanBeCanceled || operation.CancellationToken == methodToken) return null;
            if (!operation.CancellationToken.CanBeCanceled) {
                operation.CancellationToken = methodToken;
                return null;
            }
            var linked = CancellationTokenSource.CreateLinkedTokenSource(operation.CancellationToken, methodToken);
            operation.CancellationToken = linked.Token;
            return linked;
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
