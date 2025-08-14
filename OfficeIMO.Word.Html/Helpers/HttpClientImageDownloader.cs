using System;
using System.Collections.Concurrent;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Html.Helpers {
    internal class HttpClientImageDownloader : IImageDownloader {
        private readonly HttpClient _httpClient;
        private static readonly ConcurrentDictionary<string, byte[]> _cache = new(StringComparer.OrdinalIgnoreCase);

        public HttpClientImageDownloader(HttpClient? httpClient = null) {
            _httpClient = httpClient ?? new HttpClient();
        }

        public async Task<byte[]?> DownloadAsync(string src, CancellationToken cancellationToken) {
            if (_cache.TryGetValue(src, out var cached)) {
                return cached;
            }

            using var response = await _httpClient.GetAsync(src, cancellationToken).ConfigureAwait(false);
            response.EnsureSuccessStatusCode();
            var bytes = await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
            _cache[src] = bytes;
            return bytes;
        }
    }
}
