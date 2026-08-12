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
        public static ExcelWorkbookDataReader OpenDataReader(string path, ExcelReadOptions? options = null) {
            if (string.IsNullOrWhiteSpace(path)) {
                throw new ArgumentException("File path cannot be empty.", nameof(path));
            }

            ExcelReadOptions effectiveOptions = options ?? new ExcelReadOptions();
            return Path.GetExtension(path).ToLowerInvariant() switch {
                ".xls" => ExcelWorkbookDataReader.OpenLegacy(path, effectiveOptions),
                ".xlsb" => ExcelWorkbookDataReader.OpenBinary(path, effectiveOptions),
                ".xlsx" or ".xlsm" or ".xltx" or ".xltm" or ".xlam" =>
                    ExcelWorkbookDataReader.OpenOpenXml(path, effectiveOptions),
                _ => throw new NotSupportedException(
                    "OpenDataReader supports .xlsx, .xlsm, .xltx, .xltm, .xlam, .xlsb, and .xls workbooks.")
            };
        }

        /// <summary>
        /// Opens an XLSX, XLSM, XLTX, XLTM, XLAM, XLSB, or BIFF8 XLS workbook stream as a forward-only data reader.
        /// The format is detected from the package rather than the file name. The input stream
        /// remains open after the returned reader is disposed.
        /// </summary>
        /// <param name="stream">Readable workbook stream positioned at the workbook bytes to read.</param>
        /// <param name="options">Workbook reader options.</param>
        /// <returns>A forward-only reader whose results follow workbook worksheet order.</returns>
        public static ExcelWorkbookDataReader OpenDataReader(Stream stream, ExcelReadOptions? options = null) {
            if (stream == null) {
                throw new ArgumentNullException(nameof(stream));
            }
            if (!stream.CanRead) {
                throw new ArgumentException("Stream must be readable.", nameof(stream));
            }

            ExcelReadOptions effectiveOptions = options ?? new ExcelReadOptions();
            long originalPosition = stream.CanSeek ? stream.Position : 0L;
            byte[] bytes;
            try {
                bytes = OfficeIMO.Core.Internal.OfficeStreamReader.ReadRemainingBytes(
                    stream,
                    effectiveOptions.CancellationToken,
                    effectiveOptions.MaxInputBytes);
            } finally {
                if (stream.CanSeek) {
                    stream.Position = originalPosition;
                }
            }

            return OpenDataReader(bytes, effectiveOptions);
        }

        /// <summary>
        /// Opens an in-memory XLSX, XLSM, XLTX, XLTM, XLAM, XLSB, or BIFF8 XLS workbook as a forward-only data reader.
        /// </summary>
        public static ExcelWorkbookDataReader OpenDataReader(byte[] bytes, ExcelReadOptions? options = null) {
            if (bytes == null) {
                throw new ArgumentNullException(nameof(bytes));
            }

            ExcelReadOptions effectiveOptions = options ?? new ExcelReadOptions();
            effectiveOptions.CancellationToken.ThrowIfCancellationRequested();
            if (bytes.LongLength > effectiveOptions.MaxInputBytes) {
                throw new InvalidDataException(
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
        public ExcelWorkbookDataReader CreateDataReader(ExcelReadOptions? options = null) {
            ExcelReadOptions effectiveOptions = options ?? new ExcelReadOptions();
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
        public static async Task<ExcelWorkbookDataReader> OpenDataReaderAsync(
            Uri uri,
            ExcelReadOptions? options = null,
            ExcelHttpLoadOptions? httpOptions = null,
            CancellationToken cancellationToken = default) {
            if (uri == null) throw new ArgumentNullException(nameof(uri));
            ExcelReadOptions effectiveOptions = options ?? new ExcelReadOptions();
            var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                cancellationToken,
                effectiveOptions.CancellationToken);
            try {
                byte[] bytes = await ExcelHttpWorkbookLoader.DownloadAsync(
                    uri,
                    httpOptions,
                    linkedCancellation.Token,
                    effectiveOptions.MaxInputBytes).ConfigureAwait(false);
                return OpenDataReader(bytes, effectiveOptions.WithCancellationToken(linkedCancellation.Token))
                    .OwnLifetime(linkedCancellation);
            } catch {
                linkedCancellation.Dispose();
                throw;
            }
        }
    }
}
