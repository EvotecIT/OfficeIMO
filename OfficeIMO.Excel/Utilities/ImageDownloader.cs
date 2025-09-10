using System;
using System.Collections.Concurrent;
using System.IO;
using System.Net;
#if !NETFRAMEWORK
using System.Net.Http;
#endif

namespace OfficeIMO.Excel
{
    internal static class ImageDownloader
    {
        private static readonly ConcurrentDictionary<string, byte[]> Cache = new(StringComparer.OrdinalIgnoreCase);

        public static bool TryFetch(string url, int timeoutSeconds, long maxBytes, out byte[]? bytes, out string? contentType)
        {
            bytes = null; contentType = null;
            try
            {
                if (Cache.TryGetValue(url, out var cached)) { bytes = cached; return true; }
#if NETFRAMEWORK
                var request = (HttpWebRequest)WebRequest.Create(url);
                request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                request.Timeout = Math.Max(1, timeoutSeconds) * 1000;
                using (var response = (HttpWebResponse)request.GetResponse())
                {
                    if (response.StatusCode != HttpStatusCode.OK) return false;
                    var ct = response.ContentType ?? string.Empty;
                    if (!ct.StartsWith("image/", StringComparison.OrdinalIgnoreCase)) return false;
                    var len = response.ContentLength;
                    if (len > 0 && len > maxBytes) return false;
                    using var s = response.GetResponseStream();
                    if (s == null) return false;
                    using var ms = new MemoryStream(); s.CopyTo(ms);
                    if (ms.Length > maxBytes) return false;
                    var arr = ms.ToArray();
                    Cache[url] = arr;
                    bytes = arr; contentType = ct;
                    return true;
                }
#else
                using (var handler = new HttpClientHandler { AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate })
                using (var http = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(Math.Max(1, timeoutSeconds)) })
                using (var resp = http.GetAsync(url, HttpCompletionOption.ResponseHeadersRead).GetAwaiter().GetResult())
                {
                    if (!resp.IsSuccessStatusCode) return false;
                    var ct = resp.Content.Headers.ContentType?.MediaType ?? string.Empty;
                    if (!ct.StartsWith("image/", StringComparison.OrdinalIgnoreCase)) return false;
                    var len = resp.Content.Headers.ContentLength;
                    if (len.HasValue && len.Value > maxBytes) return false;
                    using var s = resp.Content.ReadAsStreamAsync().GetAwaiter().GetResult();
                    using var ms = new MemoryStream(); s.CopyTo(ms);
                    if (ms.Length > maxBytes) return false;
                    var arr = ms.ToArray();
                    Cache[url] = arr;
                    bytes = arr; contentType = ct;
                    return true;
                }
#endif
            }
            catch { return false; }
        }
    }
}
