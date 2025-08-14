using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Html.Helpers {
    /// <summary>
    /// Abstraction for downloading external images.
    /// </summary>
    public interface IImageDownloader {
        /// <summary>
        /// Downloads image data from the specified source.
        /// </summary>
        /// <param name="src">Image source URI.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        /// <returns>Image bytes or <c>null</c> if download failed.</returns>
        Task<byte[]?> DownloadAsync(string src, CancellationToken cancellationToken);
    }
}
