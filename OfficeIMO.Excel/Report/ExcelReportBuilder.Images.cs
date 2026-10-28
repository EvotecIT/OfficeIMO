using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    // Image authoring forwards to the sheet and returns its attached image.
    public sealed partial class ExcelReportBuilder {
        /// <summary>
        /// Inserts an image anchored to a specific cell.
        /// </summary>
        /// <param name="row">1-based row index.</param>
        /// <param name="column">1-based column index.</param>
        /// <param name="bytes">Image bytes.</param>
        /// <param name="contentType">Image content type (e.g., image/png).</param>
        /// <param name="width">Width as a document length.</param>
        /// <param name="height">Height as a document length.</param>
        /// <param name="offsetX">Optional X offset as a document length.</param>
        /// <param name="offsetY">Optional Y offset as a document length.</param>
        public ExcelImage AddImage(int row, int column, byte[] bytes, string contentType = "image/png", OfficeLength? width = null, OfficeLength? height = null, OfficeLength offsetX = default, OfficeLength offsetY = default) {
            return Sheet.AddImage(row, column, bytes, contentType, width, height, offsetX, offsetY);
        }

        /// <summary>
        /// Asynchronously downloads an image from a URL and inserts it anchored at a specific cell.
        /// </summary>
        /// <param name="row">1-based row index.</param>
        /// <param name="column">1-based column index.</param>
        /// <param name="url">Direct URL to an image.</param>
        /// <param name="width">Width as a document length.</param>
        /// <param name="height">Height as a document length.</param>
        /// <param name="offsetX">Optional X offset as a document length.</param>
        /// <param name="offsetY">Optional Y offset as a document length.</param>
        /// <param name="cancellationToken">Token used to cancel the remote request.</param>
        public async Task<ExcelImage> AddImageFromUrlAsync(int row, int column, string url, OfficeLength? width = null,
            OfficeLength? height = null, OfficeLength offsetX = default, OfficeLength offsetY = default,
            CancellationToken cancellationToken = default) {
            return await Sheet.AddImageFromUrlAsync(row, column, url, width, height, offsetX: offsetX,
                offsetY: offsetY, cancellationToken: cancellationToken).ConfigureAwait(false);
        }

        /// <summary>Downloads an image with an explicit remote-network policy and inserts it at a cell.</summary>
        public async Task<ExcelImage> AddImageFromUrlAsync(int row, int column, string url,
            OfficeRemoteImageLoadOptions remoteImageOptions, OfficeLength? width = null, OfficeLength? height = null,
            OfficeLength offsetX = default, OfficeLength offsetY = default, CancellationToken cancellationToken = default) {
            if (remoteImageOptions == null) throw new ArgumentNullException(nameof(remoteImageOptions));
            return await Sheet.AddImageFromUrlWithOptionsAsync(row, column, url, remoteImageOptions, width, height,
                offsetX, offsetY, cancellationToken: cancellationToken).ConfigureAwait(false);
        }
    }
}
