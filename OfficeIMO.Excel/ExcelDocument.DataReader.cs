#nullable enable

using System.Data.Common;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Opens an XLSX, XLSM, XLTX, XLTM, XLAM, XLSB, or BIFF8 XLS workbook as a forward-only data reader.
        /// Worksheets are exposed in workbook order through <see cref="DbDataReader.NextResult"/>.
        /// </summary>
        public static ExcelWorkbookDataReader OpenReader(string path, ExcelReaderOptions? options = null) {
            if (string.IsNullOrWhiteSpace(path)) {
                throw new ArgumentException("File path cannot be empty.", nameof(path));
            }

            ExcelReaderOptions effectiveOptions = options ?? new ExcelReaderOptions();
            if (effectiveOptions.Password is string password) return OpenEncryptedDataReader(path, password, effectiveOptions);
            return Path.GetExtension(path).ToLowerInvariant() switch {
                ".xls" => ExcelWorkbookDataReader.OpenLegacy(path, effectiveOptions),
                ".xlsb" => ExcelWorkbookDataReader.OpenBinary(path, effectiveOptions),
                ".xlsx" or ".xlsm" or ".xltx" or ".xltm" or ".xlam" =>
                    ExcelWorkbookDataReader.OpenOpenXml(path, effectiveOptions),
                _ => throw new NotSupportedException(
                    "OpenReader supports .xlsx, .xlsm, .xltx, .xltm, .xlam, .xlsb, and .xls workbooks.")
            };
        }

        /// <summary>
        /// Reads worksheet names from XLSX, XLSM, XLTX, XLTM, XLAM, XLSB, or BIFF5/BIFF8 XLS
        /// workbook metadata without opening a worksheet data stream.
        /// </summary>
        /// <param name="path">Workbook path.</param>
        /// <param name="options">Workbook metadata read limits and cancellation.</param>
        /// <returns>Readable worksheet names in workbook order.</returns>
        internal static IReadOnlyList<string> GetSheetNames(string path, ExcelReaderOptions? options = null) {
            if (string.IsNullOrWhiteSpace(path)) {
                throw new ArgumentException("File path cannot be empty.", nameof(path));
            }

            ExcelReaderOptions effectiveOptions = options ?? new ExcelReaderOptions();
            return ExcelWorkbookDataReader.ReadSheetNames(path, effectiveOptions);
        }

        /// <summary>
        /// Opens an XLSX, XLSM, XLTX, XLTM, XLAM, XLSB, or BIFF8 XLS workbook stream as a forward-only data reader.
        /// The format is detected from the package rather than the file name. The input stream
        /// remains open after the returned reader is disposed.
        /// </summary>
        /// <param name="stream">Readable workbook stream positioned at the workbook bytes to read.</param>
        /// <param name="options">Workbook reader options.</param>
        /// <returns>A forward-only reader whose results follow workbook worksheet order.</returns>
        public static ExcelWorkbookDataReader OpenReader(Stream stream, ExcelReaderOptions? options = null) {
            if (stream == null) {
                throw new ArgumentNullException(nameof(stream));
            }
            if (!stream.CanRead) {
                throw new ArgumentException("Stream must be readable.", nameof(stream));
            }

            ExcelReaderOptions effectiveOptions = options ?? new ExcelReaderOptions();
            if (effectiveOptions.Password is string password) return OpenEncryptedDataReader(stream, password, effectiveOptions);
            long originalPosition = stream.CanSeek ? stream.Position : 0L;
            byte[] bytes;
            try {
                bytes = ExcelReadLimitFailure.ReadInput(() => OfficeIMO.Core.Internal.OfficeStreamReader.ReadRemainingBytes(
                    stream,
                    effectiveOptions.CancellationToken,
                    effectiveOptions.MaxInputBytes), effectiveOptions.MaxInputBytes);
            } finally {
                if (stream.CanSeek) {
                    stream.Position = originalPosition;
                }
            }

            return OpenReader(bytes, effectiveOptions);
        }

        /// <summary>
        /// Opens an in-memory XLSX, XLSM, XLTX, XLTM, XLAM, XLSB, or BIFF8 XLS workbook as a forward-only data reader.
        /// </summary>
        internal static ExcelWorkbookDataReader OpenReader(byte[] bytes, ExcelReaderOptions? options = null) {
            if (bytes == null) {
                throw new ArgumentNullException(nameof(bytes));
            }

            ExcelReaderOptions effectiveOptions = options ?? new ExcelReaderOptions();
            effectiveOptions.CancellationToken.ThrowIfCancellationRequested();
            if (effectiveOptions.Password is string password) return OpenEncryptedDataReader(bytes, password, effectiveOptions);
            if (bytes.LongLength > effectiveOptions.MaxInputBytes) {
                throw ExcelReadLimitFailure.Create(nameof(ExcelReaderOptions.MaxInputBytes), effectiveOptions.MaxInputBytes,
                    $"Workbook input contains {bytes.LongLength} bytes, exceeding the configured limit of {effectiveOptions.MaxInputBytes} bytes.");
            }

            return ExcelDocumentLoadRouting.DetectFormat(bytes, filePath: null) switch {
                ExcelFileFormat.Xls => ExcelWorkbookDataReader.OpenLegacy(bytes, effectiveOptions),
                ExcelFileFormat.Xlsb => ExcelWorkbookDataReader.OpenBinary(bytes, effectiveOptions),
                _ => ExcelWorkbookDataReader.OpenOpenXml(bytes, effectiveOptions)
            };
        }

        /// <summary>
        /// Creates a forward-only data reader over this open workbook document.
        /// The returned reader does not close the document.
        /// </summary>
        public ExcelWorkbookDataReader CreateDataReader(ExcelReaderOptions? options = null) {
            ExcelReaderOptions effectiveOptions = options ?? new ExcelReaderOptions();
            effectiveOptions.CancellationToken.ThrowIfCancellationRequested();
            if (_spreadSheetDocument is null) {
                throw new ObjectDisposedException(nameof(ExcelDocument));
            }
            MaterializeDeferredDataSetImport(effectiveOptions.CancellationToken);
            return ExcelWorkbookDataReader.WrapOpenXml(
                ExcelDocumentReader.Wrap(_spreadSheetDocument, effectiveOptions),
                effectiveOptions);
        }

        /// <summary>
        /// Downloads an HTTP or HTTPS workbook and opens it as a forward-only data reader.
        /// </summary>
        public static async Task<ExcelWorkbookDataReader> OpenReaderAsync(
            Uri uri,
            ExcelReaderOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (uri == null) throw new ArgumentNullException(nameof(uri));
            ExcelReaderOptions effectiveOptions = (options ?? new ExcelReaderOptions()).Clone();
            var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                cancellationToken,
                effectiveOptions.CancellationToken);
            try {
                byte[] bytes = await ExcelHttpWorkbookLoader.DownloadAsync(
                    uri,
                    effectiveOptions.Transport,
                    linkedCancellation.Token,
                    effectiveOptions.MaxInputBytes).ConfigureAwait(false);
                return OpenReader(bytes, effectiveOptions.WithCancellationToken(linkedCancellation.Token))
                    .OwnLifetime(linkedCancellation);
            } catch {
                linkedCancellation.Dispose();
                throw;
            }
        }
    }
}
