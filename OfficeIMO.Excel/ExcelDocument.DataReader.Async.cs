#nullable enable

using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Asynchronously reads a local XLSX, XLSM, XLTX, XLTM, XLAM, XLSB, or BIFF8 XLS file
        /// and opens it as a forward-only data reader.
        /// </summary>
        /// <param name="path">Workbook path with a supported Excel extension.</param>
        /// <param name="options">Worksheet selection, reader limits, conversion, and cancellation settings.</param>
        /// <param name="cancellationToken">Cancels opening and remains active for the returned reader.</param>
        /// <returns>A reader owned by the caller. The input file is closed before this task completes.</returns>
        /// <remarks>
        /// Input is buffered up to <see cref="ExcelReadOptions.MaxInputBytes"/> using asynchronous I/O.
        /// Workbook validation and worksheet discovery run synchronously after input is read.
        /// Options are copied before asynchronous input begins.
        /// Both cancellation tokens remain active until the returned reader is disposed.
        /// </remarks>
        public static async Task<ExcelWorkbookDataReader> OpenDataReaderAsync(
            string path,
            ExcelReadOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(path)) {
                throw new ArgumentException("File path cannot be empty.", nameof(path));
            }

            ExcelFileFormat format = Path.GetExtension(path).ToLowerInvariant() switch {
                ".xls" => ExcelFileFormat.Xls,
                ".xlsb" => ExcelFileFormat.Xlsb,
                ".xlsx" or ".xlsm" or ".xltx" or ".xltm" or ".xlam" => ExcelFileFormat.Xlsx,
                _ => throw new NotSupportedException(
                    "OpenDataReader supports .xlsx, .xlsm, .xltx, .xltm, .xlam, .xlsb, and .xls workbooks.")
            };
            cancellationToken.ThrowIfCancellationRequested();
            options?.CancellationToken.ThrowIfCancellationRequested();

            using var source = new FileStream(
                path, FileMode.Open, FileAccess.Read, FileShare.Read, 81920,
                FileOptions.Asynchronous | FileOptions.SequentialScan);
            return await OpenLocalDataReaderAsync(
                source, options, cancellationToken, format).ConfigureAwait(false);
        }

        /// <summary>
        /// Asynchronously reads the remaining workbook bytes from a stream and opens a forward-only data reader.
        /// The physical XLSX, XLSM, XLTX, XLTM, XLAM, XLSB, or BIFF8 XLS format is detected from those bytes.
        /// </summary>
        /// <param name="stream">Readable stream positioned at the workbook bytes to read.</param>
        /// <param name="options">Worksheet selection, reader limits, conversion, and cancellation settings.</param>
        /// <param name="cancellationToken">Cancels opening and remains active for the returned reader.</param>
        /// <returns>A reader owned by the caller.</returns>
        /// <remarks>
        /// The source remains open on success, failure, and reader disposal. Seekable source positions are
        /// restored after reading; non-seekable sources advance through the remaining content.
        /// Input is buffered up to <see cref="ExcelReadOptions.MaxInputBytes"/> using asynchronous I/O.
        /// Workbook validation and worksheet discovery run synchronously after input is read.
        /// Options are copied before asynchronous input begins.
        /// Both cancellation tokens remain active until the returned reader is disposed.
        /// </remarks>
        public static Task<ExcelWorkbookDataReader> OpenDataReaderAsync(
            Stream stream,
            ExcelReadOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
            return OpenLocalDataReaderAsync(stream, options, cancellationToken, format: null);
        }

        private static async Task<ExcelWorkbookDataReader> OpenLocalDataReaderAsync(
            Stream source,
            ExcelReadOptions? options,
            CancellationToken cancellationToken,
            ExcelFileFormat? format) {
            ExcelReadOptions effectiveOptions = (options ?? new ExcelReadOptions()).Clone();
            var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                cancellationToken, effectiveOptions.CancellationToken);
            try {
                long originalPosition = source.CanSeek ? source.Position : 0L;
                byte[] bytes;
                try {
                    bytes = await OfficeIMO.Core.Internal.OfficeStreamReader.ReadRemainingBytesAsync(
                        source, linkedCancellation.Token, effectiveOptions.MaxInputBytes).ConfigureAwait(false);
                } finally {
                    if (source.CanSeek) source.Position = originalPosition;
                }

                linkedCancellation.Token.ThrowIfCancellationRequested();
                ExcelReadOptions readerOptions = effectiveOptions.WithCancellationToken(linkedCancellation.Token);
                ExcelWorkbookDataReader reader = format switch {
                    ExcelFileFormat.Xls => ExcelWorkbookDataReader.OpenLegacy(bytes, readerOptions),
                    ExcelFileFormat.Xlsb => ExcelWorkbookDataReader.OpenBinary(bytes, readerOptions),
                    ExcelFileFormat.Xlsx => ExcelWorkbookDataReader.OpenOpenXml(bytes, readerOptions),
                    _ => OpenDataReader(bytes, readerOptions)
                };
                return reader.OwnLifetime(linkedCancellation);
            } catch {
                linkedCancellation.Dispose();
                throw;
            }
        }
    }
}
