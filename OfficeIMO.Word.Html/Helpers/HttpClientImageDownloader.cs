using System;
using System.Collections.Concurrent;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Html.Helpers {
    /// <summary>
    /// Default image downloader using <see cref="HttpClient"/> with in-memory caching.
    /// </summary>
    internal class HttpClientImageDownloader : IImageDownloader, IDisposable {
        private readonly HttpClient _client = new();
        private readonly ConcurrentDictionary<string, byte[]> _cache = new(StringComparer.OrdinalIgnoreCase);
        private bool _disposed;

        public async Task<byte[]> DownloadAsync(string uri, CancellationToken cancellationToken) {
            if (_cache.TryGetValue(uri, out var bytes)) {
                return bytes;
            }

            var data = await _client.GetByteArrayAsync(uri, cancellationToken).ConfigureAwait(false);
            _cache[uri] = data;
            return data;
        }

        public void Dispose() {
            if (!_disposed) {
                _client.Dispose();
                _disposed = true;
            }
        }
    }
}
