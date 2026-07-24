using System;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelDocumentReader {
        /// <summary>
        /// Asynchronously opens a remote Excel workbook for read-only access.
        /// </summary>
        /// <param name="uri">HTTP or HTTPS URI of the workbook.</param>
        /// <param name="options">Optional read options.</param>
        /// <param name="httpOptions">Optional HTTP loading options.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>Workbook reader.</returns>
        public static async Task<ExcelDocumentReader> OpenAsync(Uri uri, ExcelReadOptions? options = null, ExcelHttpLoadOptions? httpOptions = null, CancellationToken cancellationToken = default) {
            var effectiveOptions = options ?? new ExcelReadOptions();
            byte[] bytes = await ExcelHttpWorkbookLoader.DownloadAsync(
                uri,
                httpOptions,
                cancellationToken,
                effectiveOptions.MaxInputBytes).ConfigureAwait(false);
            return OpenFromBytes(
                bytes,
                effectiveOptions,
                normalizeContentTypes: false,
                contextMessage: $"Failed to open remote workbook '{uri}' after normalizing package content types. The package may declare an invalid content type for '/docProps/app.xml'.");
        }
    }
}
