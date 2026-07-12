using DocumentFormat.OpenXml.Packaging;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Loads an Excel workbook from a remote URI.
        /// </summary>
        /// <param name="uri">HTTP or HTTPS URI of the workbook.</param>
        /// <param name="httpOptions">Optional HTTP loading options.</param>
        /// <param name="options">Access, persistence, and low-level package options.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        public static ExcelDocument Load(Uri uri, ExcelHttpLoadOptions? httpOptions = null, ExcelLoadOptions? options = null) {
            byte[] bytes = ExcelHttpWorkbookLoader.Download(uri, httpOptions, CancellationToken.None);
            return LoadFromByteArray(bytes, options ?? new ExcelLoadOptions(), filePath: null);
        }

        /// <summary>
        /// Asynchronously loads an Excel workbook from a remote URI.
        /// </summary>
        /// <param name="uri">HTTP or HTTPS URI of the workbook.</param>
        /// <param name="httpOptions">Optional HTTP loading options.</param>
        /// <param name="options">Access, persistence, and low-level package options.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        public static async Task<ExcelDocument> LoadAsync(Uri uri, ExcelHttpLoadOptions? httpOptions = null, ExcelLoadOptions? options = null, CancellationToken cancellationToken = default) {
            byte[] bytes = await ExcelHttpWorkbookLoader.DownloadAsync(uri, httpOptions, cancellationToken).ConfigureAwait(false);
            return LoadFromByteArray(bytes, options ?? new ExcelLoadOptions(), filePath: null);
        }
    }
}
