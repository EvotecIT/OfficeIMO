using DocumentFormat.OpenXml.Packaging;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Asynchronously loads an Excel workbook from a remote URI.
        /// </summary>
        /// <param name="uri">HTTP or HTTPS URI of the workbook.</param>
        /// <param name="options">Access, persistence, and low-level package options.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        public static async Task<ExcelDocument> LoadAsync(Uri uri, ExcelLoadOptions? options = null, CancellationToken cancellationToken = default) {
            ExcelLoadOptions resolved = (options ?? new ExcelLoadOptions()).Snapshot();
            ValidateRemoteLoadLifecycle(resolved);
            long? inputLimit = ResolveInputLimit(resolved);
            bool inputLimitIsPackageSize = resolved.PackageSecurity != null
                && inputLimit == resolved.PackageSecurity.MaxPackageBytes;
            byte[] bytes = await ExcelHttpWorkbookLoader.DownloadAsync(
                uri,
                resolved,
                cancellationToken,
                inputLimit,
                sourceLimitIsPackageSize: inputLimitIsPackageSize).ConfigureAwait(false);
            return LoadFromByteArray(bytes, resolved, filePath: null);
        }

        private static void ValidateRemoteLoadLifecycle(ExcelLoadOptions options) {
            OfficeIMO.Core.Internal.OfficeDocumentLifecycle.Validate(
                options.AccessMode,
                options.PersistenceMode,
                "workbook");
            if (options.PersistenceMode == OfficeIMO.DocumentPersistenceMode.SaveOnDispose) {
                throw new ArgumentException(
                    "SaveOnDispose requires an associated file path or writable stream. Remote URI loads are detached and must be saved explicitly.",
                    nameof(options));
            }
        }
    }
}
