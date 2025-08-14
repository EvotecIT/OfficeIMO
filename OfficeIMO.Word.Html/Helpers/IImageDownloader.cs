using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Html.Helpers {
    /// <summary>
    /// Downloads image data for the HTML to Word converter.
    /// Implementations may provide caching or custom retrieval logic.
    /// </summary>
    public interface IImageDownloader {
        /// <summary>
        /// Retrieves the raw bytes for the specified URI.
        /// </summary>
        /// <param name="uri">Image location.</param>
        /// <param name="cancellationToken">Token used to cancel the request.</param>
        /// <returns>Byte array of the downloaded content.</returns>
        Task<byte[]> DownloadAsync(string uri, CancellationToken cancellationToken);
    }
}
