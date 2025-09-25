using System.Collections.Concurrent;
using System.Net;
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

        public static bool TryFetch(string url, int timeoutSeconds, long maxBytes, out byte[]? bytes, out string? contentType) {
            bytes = null; contentType = null;
            try {
                if (Cache.TryGetValue(url, out var cached)) {
                    if (DateTimeOffset.UtcNow <= cached.ExpiresAt) {
                        bytes = cached.Bytes;
                        contentType = cached.ContentType;
                        return true;
                    }

                    Cache.TryRemove(url, out _);
                }
#if NETFRAMEWORK
                var request = (HttpWebRequest)WebRequest.Create(url);
                request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                request.Timeout = Math.Max(1, timeoutSeconds) * 1000;
                using (var response = (HttpWebResponse)request.GetResponse())
                {
                    if (response.StatusCode != HttpStatusCode.OK) return false;
                    var ctRaw = response.ContentType;
                    var ct = string.IsNullOrWhiteSpace(ctRaw) ? null : ctRaw;
                    if (ct == null || !ct.StartsWith("image/", StringComparison.OrdinalIgnoreCase)) return false;
                    var len = response.ContentLength;
                    if (len > 0 && len > maxBytes) return false;
                    using var s = response.GetResponseStream();
                    if (s == null) return false;
                    using var ms = new MemoryStream(); s.CopyTo(ms);
                    if (ms.Length > maxBytes) return false;
                    var arr = ms.ToArray();
                    Cache[url] = new CacheEntry(arr, ct, DateTimeOffset.UtcNow.Add(CacheEntryLifetime));
                    CacheOrder.Enqueue(url);
                    TrimCache();
                    bytes = arr; contentType = ct;
                    return true;
                }
#else
                using (var handler = new HttpClientHandler { AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate })
                using (var http = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(Math.Max(1, timeoutSeconds)) })
                using (var resp = http.GetAsync(url, HttpCompletionOption.ResponseHeadersRead).GetAwaiter().GetResult()) {
                    if (!resp.IsSuccessStatusCode) return false;
                    var ctRaw = resp.Content.Headers.ContentType?.MediaType;
                    var ct = string.IsNullOrWhiteSpace(ctRaw) ? null : ctRaw;
                    if (ct == null || !ct.StartsWith("image/", StringComparison.OrdinalIgnoreCase)) return false;
                    var len = resp.Content.Headers.ContentLength;
                    if (len.HasValue && len.Value > maxBytes) return false;
                    using var s = resp.Content.ReadAsStreamAsync().GetAwaiter().GetResult();
                    using var ms = new MemoryStream(); s.CopyTo(ms);
                    if (ms.Length > maxBytes) return false;
                    var arr = ms.ToArray();
                    Cache[url] = new CacheEntry(arr, ct, DateTimeOffset.UtcNow.Add(CacheEntryLifetime));
                    CacheOrder.Enqueue(url);
                    TrimCache();
                    bytes = arr; contentType = ct;
                    return true;
                }
#endif
            } catch { return false; }
        }
    }
}
