#nullable enable

using OfficeIMO.Core.Internal;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>Asynchronously reads and decrypts a local Office Open XML workbook and opens a forward-only data reader.</summary>
        /// <param name="path">Path to an encrypted XLSX, XLSM, XLTX, XLTM, XLAM, or XLSB package.</param>
        /// <param name="password">The package password.</param>
        /// <param name="options">Worksheet selection, reader limits, conversion, and cancellation settings.</param>
        /// <param name="cancellationToken">Cancels opening and remains active for the returned reader.</param>
        /// <returns>A reader owned by the caller. The input file is closed before the task completes.</returns>
        /// <remarks>
        /// Asynchronous input is followed by synchronous decryption, integrity verification and workbook discovery.
        /// Encrypted input and decrypted package sizes are each bounded by ExcelReadOptions.MaxInputBytes.
        /// Opening buffers both complete packages and copies options before input begins. Both cancellation
        /// tokens remain active until the reader is disposed. Only Office Agile encryption is supported.
        /// Legacy XLS password encryption uses LoadEncrypted.
        /// </remarks>
        public static async Task<ExcelWorkbookDataReader> OpenEncryptedDataReaderAsync(
            string path, string password, ExcelReadOptions? options = null, CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
            if (password == null) throw new ArgumentNullException(nameof(password));
            ExcelReadOptions effectiveOptions = (options ?? new ExcelReadOptions()).Clone();
            cancellationToken.ThrowIfCancellationRequested();
            effectiveOptions.CancellationToken.ThrowIfCancellationRequested();
            using var source = new FileStream(
                path, FileMode.Open, FileAccess.Read, FileShare.Read, 81920,
                FileOptions.Asynchronous | FileOptions.SequentialScan);
            return await OpenEncryptedDataReaderAsyncCore(
                source, password, effectiveOptions, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>Asynchronously reads and decrypts the remaining Office Open XML workbook bytes and opens a forward-only reader.</summary>
        /// <param name="stream">Readable stream positioned at the encrypted workbook bytes.</param>
        /// <param name="password">The package password.</param>
        /// <param name="options">Worksheet selection, reader limits, conversion, and cancellation settings.</param>
        /// <param name="cancellationToken">Cancels opening and remains active for the returned reader.</param>
        /// <returns>A reader owned by the caller.</returns>
        /// <remarks>
        /// The input remains open on success, failure, and reader disposal. Seekable positions are restored;
        /// non-seekable sources advance through their remaining content. Both complete packages are buffered
        /// and each is bounded by ExcelReadOptions.MaxInputBytes. Options are copied before asynchronous input.
        /// Decryption, integrity verification and workbook discovery run synchronously after input is read.
        /// Both cancellation tokens remain active until reader disposal. This API supports Office Open XML packages protected by Agile encryption.
        /// </remarks>
        public static Task<ExcelWorkbookDataReader> OpenEncryptedDataReaderAsync(
            Stream stream, string password, ExcelReadOptions? options = null, CancellationToken cancellationToken = default) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
            if (password == null) throw new ArgumentNullException(nameof(password));
            return OpenEncryptedDataReaderAsyncCore(
                stream, password, (options ?? new ExcelReadOptions()).Clone(), cancellationToken);
        }

        private static async Task<ExcelWorkbookDataReader> OpenEncryptedDataReaderAsyncCore(
            Stream source, string password, ExcelReadOptions options, CancellationToken cancellationToken) {
            var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                cancellationToken, options.CancellationToken);
            try {
                long originalPosition = source.CanSeek ? source.Position : 0L;
                byte[] encryptedBytes;
                try {
                    encryptedBytes = await OfficeStreamReader.ReadRemainingBytesAsync(
                        source, linkedCancellation.Token, options.MaxInputBytes).ConfigureAwait(false);
                } finally {
                    if (source.CanSeek) source.Position = originalPosition;
                }
                return OpenEncryptedDataReaderCore(
                    encryptedBytes, password, options.WithCancellationToken(linkedCancellation.Token))
                    .OwnLifetime(linkedCancellation);
            } catch {
                linkedCancellation.Dispose();
                throw;
            }
        }
    }
}
