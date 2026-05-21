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
        /// <param name="readOnly">Open the document in read-only mode. Remote loads are read-only by default.</param>
        /// <param name="openSettings">Optional Open XML settings to control how the package is opened.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        public static ExcelDocument Load(Uri uri, ExcelHttpLoadOptions? httpOptions = null, bool readOnly = true, OpenSettings? openSettings = null) {
            byte[] bytes = ExcelHttpWorkbookLoader.Download(uri, httpOptions, CancellationToken.None);
            return LoadFromByteArray(bytes, readOnly, autoSave: false, filePath: null, log: null, openSettings, preferFilePathOnFallback: false);
        }

        /// <summary>
        /// Asynchronously loads an Excel workbook from a remote URI.
        /// </summary>
        /// <param name="uri">HTTP or HTTPS URI of the workbook.</param>
        /// <param name="httpOptions">Optional HTTP loading options.</param>
        /// <param name="readOnly">Open the document in read-only mode. Remote loads are read-only by default.</param>
        /// <param name="openSettings">Optional Open XML settings to control how the package is opened.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        public static async Task<ExcelDocument> LoadAsync(Uri uri, ExcelHttpLoadOptions? httpOptions = null, bool readOnly = true, OpenSettings? openSettings = null, CancellationToken cancellationToken = default) {
            byte[] bytes = await ExcelHttpWorkbookLoader.DownloadAsync(uri, httpOptions, cancellationToken).ConfigureAwait(false);
            return LoadFromByteArray(bytes, readOnly, autoSave: false, filePath: null, log: null, openSettings, preferFilePathOnFallback: false);
        }
    }
}
