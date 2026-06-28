using System.Collections.Concurrent;
using System.Net;
using OfficeIMO.Drawing;
#if !NETFRAMEWORK
using System.Net.Http;
#endif

namespace OfficeIMO.Excel {
    internal static class ImageDownloader {
        private sealed class CacheEntry {
            public CacheEntry(byte[] bytes, string? contentType, DateTimeOffset expiresAt) {
                Bytes = bytes;
                ContentType = contentType;
                ExpiresAt = expiresAt;
            }

            public byte[] Bytes { get; }
            public string? ContentType { get; }
            public DateTimeOffset ExpiresAt { get; }
        }

        private const int CacheCapacity = 32;
        private const int MaxRedirects = 5;
        private const int BufferSize = 81920;
        private static readonly TimeSpan CacheEntryLifetime = TimeSpan.FromMinutes(10);
        private static readonly ConcurrentDictionary<string, CacheEntry> Cache = new(StringComparer.OrdinalIgnoreCase);
        private static readonly ConcurrentQueue<string> CacheOrder = new();

        internal static void ClearCache() {
            while (CacheOrder.TryDequeue(out _)) { }
            Cache.Clear();
        }

        private static void TrimCache() {
            while (Cache.Count > CacheCapacity && CacheOrder.TryDequeue(out var oldestKey)) {
                Cache.TryRemove(oldestKey, out _);
            }
        }

        private static string? NormalizeContentType(string? raw) {
            return OfficeImageInfo.TryNormalizeImageContentType(raw, out string normalizedContentType)
                ? normalizedContentType
                : null;
        }

        public static bool TryFetch(string url, int timeoutSeconds, long maxBytes, out byte[]? bytes, out string? contentType) {
            bytes = null; contentType = null;
            try {
                if (maxBytes <= 0 || !TryCreateHttpUri(url, out var uri)) return false;

                var cacheKey = uri.AbsoluteUri;
                if (Cache.TryGetValue(cacheKey, out var cached)) {
                    if (DateTimeOffset.UtcNow <= cached.ExpiresAt) {
                        bytes = cached.Bytes;
                        contentType = cached.ContentType;
                        return true;
                    }

                    Cache.TryRemove(cacheKey, out _);
                }
#if NETFRAMEWORK
                using (var response = SendWithRedirects(uri, timeoutSeconds))
#else
                using (var handler = new HttpClientHandler { AllowAutoRedirect = false, AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate })
                using (var http = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(Math.Max(1, timeoutSeconds)) })
                using (var response = SendWithRedirects(http, uri))
#endif
                {
                    if (response == null) return false;
#if NETFRAMEWORK
                    if (response.StatusCode != HttpStatusCode.OK) return false;
                    var ct = NormalizeContentType(response.ContentType);
                    var len = response.ContentLength;
#else
                    if (!response.IsSuccessStatusCode) return false;
                    var ct = NormalizeContentType(response.Content.Headers.ContentType?.MediaType);
                    var len = response.Content.Headers.ContentLength;
#endif
                    if (ct == null) return false;
#if NETFRAMEWORK
                    if (len > 0 && len > maxBytes) return false;
                    using var s = response.GetResponseStream();
#else
                    if (len.HasValue && len.Value > maxBytes) return false;
                    using var s = response.Content.ReadAsStreamAsync().GetAwaiter().GetResult();
#endif
                    if (s == null) return false;
                    var arr = ReadWithLimit(s, maxBytes);
                    if (arr == null) return false;
                    Cache[cacheKey] = new CacheEntry(arr, ct, DateTimeOffset.UtcNow.Add(CacheEntryLifetime));
                    CacheOrder.Enqueue(cacheKey);
                    TrimCache();
                    bytes = arr; contentType = ct;
                    return true;
                }
            } catch { return false; }
        }

        private static bool TryCreateHttpUri(string url, out Uri uri) {
            uri = null!;
            if (string.IsNullOrWhiteSpace(url) || !Uri.TryCreate(url, UriKind.Absolute, out var parsed)) return false;
            if (!IsHttpUri(parsed)) return false;

            uri = parsed;
            return true;
        }

        private static bool IsHttpUri(Uri uri) {
            return string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)
                || string.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsRedirect(HttpStatusCode statusCode) {
            return statusCode == HttpStatusCode.Moved
                || statusCode == HttpStatusCode.Redirect
                || statusCode == HttpStatusCode.SeeOther
                || statusCode == HttpStatusCode.TemporaryRedirect
                || (int)statusCode == 308;
        }

        private static Uri? ResolveRedirect(Uri currentUri, string? location) {
            if (string.IsNullOrWhiteSpace(location)) return null;
            if (!Uri.TryCreate(location, UriKind.RelativeOrAbsolute, out var parsed)) return null;

            var resolved = parsed.IsAbsoluteUri ? parsed : new Uri(currentUri, parsed);
            return IsHttpUri(resolved) && IsSameOrigin(currentUri, resolved) ? resolved : null;
        }

        private static bool IsSameOrigin(Uri left, Uri right) {
            return string.Equals(left.Scheme, right.Scheme, StringComparison.OrdinalIgnoreCase)
                   && string.Equals(NormalizeHost(left), NormalizeHost(right), StringComparison.OrdinalIgnoreCase)
                   && left.Port == right.Port;
        }

        private static string NormalizeHost(Uri uri) {
            string host = string.IsNullOrEmpty(uri.IdnHost) ? uri.Host : uri.IdnHost;
            return host.TrimEnd('.').ToLowerInvariant();
        }

#if NETFRAMEWORK
        private static HttpWebResponse? SendWithRedirects(Uri uri, int timeoutSeconds) {
            var currentUri = uri;
            for (int redirectCount = 0; redirectCount <= MaxRedirects; redirectCount++) {
                var request = (HttpWebRequest)WebRequest.Create(currentUri);
                request.AllowAutoRedirect = false;
                request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                request.Timeout = Math.Max(1, timeoutSeconds) * 1000;

                var response = (HttpWebResponse)request.GetResponse();
                if (!IsRedirect(response.StatusCode)) return response;

                var nextUri = ResolveRedirect(currentUri, response.Headers[HttpResponseHeader.Location]);
                response.Dispose();
                if (nextUri == null || redirectCount == MaxRedirects) return null;
                currentUri = nextUri;
            }

            return null;
        }
#else
        private static HttpResponseMessage? SendWithRedirects(HttpClient http, Uri uri) {
            var currentUri = uri;
            for (int redirectCount = 0; redirectCount <= MaxRedirects; redirectCount++) {
                var response = http.GetAsync(currentUri, HttpCompletionOption.ResponseHeadersRead).GetAwaiter().GetResult();
                if (!IsRedirect(response.StatusCode)) return response;

                var nextUri = ResolveRedirect(currentUri, response.Headers.Location?.ToString());
                response.Dispose();
                if (nextUri == null || redirectCount == MaxRedirects) return null;
                currentUri = nextUri;
            }

            return null;
        }
#endif

        private static byte[]? ReadWithLimit(Stream stream, long maxBytes) {
            using var ms = new MemoryStream();
            var buffer = new byte[BufferSize];
            long total = 0;
            while (true) {
                int read = stream.Read(buffer, 0, buffer.Length);
                if (read == 0) break;

                total += read;
                if (total > maxBytes) return null;
                ms.Write(buffer, 0, read);
            }

            return ms.ToArray();
        }
    }
}
