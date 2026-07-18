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
        /// <returns>The generated first-party PDF document model.</returns>
        public static PdfCore.PdfDocument ToPdfDocument(this WordDocument document, PdfSaveOptions? options = null) {
            return document.ToPdfDocumentResult(options).Value;
        }

        /// <summary>
        /// Converts the specified <see cref="WordDocument"/> to a PDF document and returns conversion diagnostics with it.
        /// </summary>
        public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this WordDocument document, PdfSaveOptions? options = null) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            PdfSaveOptions operation = (options ?? new PdfSaveOptions()).CloneForConversion();
            PdfCore.PdfDocument pdf = CreateOfficeIMOPdfDocument(document, operation);
            return new PdfCore.PdfDocumentConversionResult(pdf, operation.Report);
        }

        /// <summary>
        /// Saves the specified <see cref="WordDocument"/> as a PDF at the given <paramref name="path"/>.
        /// </summary>
        /// <param name="document">The document to convert.</param>
        /// <param name="path">The output PDF file path.</param>
        /// <param name="options">Optional PDF configuration.</param>
        public static PdfCore.PdfDocumentConversionResult SaveAsPdf(this WordDocument document, string path, PdfSaveOptions? options = null) {
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

            return document.ToPdfDocumentResult(options).Save(fullPath);
        }

        /// <summary>
        /// Attempts to save the specified <see cref="WordDocument"/> as a PDF file and returns output diagnostics instead of throwing.
        /// </summary>
        public static PdfCore.PdfSaveResult TrySaveAsPdf(this WordDocument document, string path, PdfSaveOptions? options = null) {
            try {
                if (document == null) {
                    throw new ArgumentNullException(nameof(document));
                }

                return document.ToPdfDocumentResult(options).TrySave(path);
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
        public static PdfCore.PdfDocumentConversionResult SaveAsPdf(this WordDocument document, Stream stream, PdfSaveOptions? options = null) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            if (stream == null) {
                throw new ArgumentNullException(nameof(stream));
            }

            if (!stream.CanWrite) {
                throw new ArgumentException("Stream must be writable.", nameof(stream));
            }

            PdfCore.PdfDocumentConversionResult result = document.ToPdfDocumentResult(options).Save(stream);
            if (stream.CanSeek) {
                stream.Position = 0;
            }
            return result;
        }

        /// <summary>
        /// Attempts to write the specified <see cref="WordDocument"/> as a PDF to a stream and returns output diagnostics instead of throwing.
        /// </summary>
        public static PdfCore.PdfSaveResult TrySaveAsPdf(this WordDocument document, Stream stream, PdfSaveOptions? options = null) {
            try {
                if (document == null) {
                    throw new ArgumentNullException(nameof(document));
                }

                PdfCore.PdfSaveResult result = document.ToPdfDocumentResult(options).TrySave(stream);
                if (result.Succeeded && stream != null && stream.CanSeek) {
                    stream.Position = 0;
                }

                return result;
            } catch (Exception ex) {
                return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex);
            }
        }

        /// <summary>
        /// Converts the specified <see cref="WordDocument"/> to PDF bytes.
        /// </summary>
        /// <param name="document">The document to convert.</param>
        /// <param name="options">Optional PDF configuration.</param>
        /// <returns>The generated PDF as a byte array.</returns>
        /// <example><code>byte[] pdf = document.ToPdf();</code></example>
        public static byte[] ToPdf(this WordDocument document, PdfSaveOptions? options = null) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            return document.ToPdfDocument(options).ToBytes();
        }

        /// <summary>
        /// Saves the specified <see cref="WordDocument"/> as a PDF at the given <paramref name="path"/> asynchronously.
        /// </summary>
        /// <param name="document">The document to convert.</param>
        /// <param name="path">The output PDF file path.</param>
        /// <param name="options">Optional PDF configuration.</param>
        /// <param name="cancellationToken">A token to observe while waiting for the task to complete.</param>
        /// <returns>The saved PDF conversion result with diagnostics.</returns>
        public static async Task<PdfCore.PdfDocumentConversionResult> SaveAsPdfAsync(this WordDocument document, string path, PdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
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

            return await document.ToPdfDocumentResult(options).SaveAsync(fullPath, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Attempts to save the specified <see cref="WordDocument"/> as a PDF file asynchronously and returns output diagnostics instead of throwing.
        /// </summary>
        public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(this WordDocument document, string path, PdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            try {
                if (document == null) {
                    throw new ArgumentNullException(nameof(document));
                }

                return await document.ToPdfDocumentResult(options).TrySaveAsync(path, cancellationToken).ConfigureAwait(false);
            } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
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
        /// <returns>The saved PDF conversion result with diagnostics.</returns>
        public static async Task<PdfCore.PdfDocumentConversionResult> SaveAsPdfAsync(this WordDocument document, Stream stream, PdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
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

            PdfCore.PdfDocumentConversionResult result = await document.ToPdfDocumentResult(options).SaveAsync(stream, cancellationToken).ConfigureAwait(false);
            if (stream.CanSeek) {
                stream.Position = 0;
            }
            return result;
        }

        /// <summary>
        /// Attempts to write the specified <see cref="WordDocument"/> as a PDF to a stream asynchronously and returns output diagnostics instead of throwing.
        /// </summary>
        public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(this WordDocument document, Stream stream, PdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            try {
                if (document == null) {
                    throw new ArgumentNullException(nameof(document));
                }

                PdfCore.PdfSaveResult result = await document.ToPdfDocumentResult(options).TrySaveAsync(stream, cancellationToken).ConfigureAwait(false);
                if (result.Succeeded && stream != null && stream.CanSeek) {
                    stream.Position = 0;
                }

                return result;
            } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
                throw;
            } catch (Exception ex) {
                return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex);
            }
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
