using System;
using System.Collections.Concurrent;
using System.IO;
using System.Net;
#if !NET472 && !NET48
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
#if NET472 || NET48
                var req = (HttpWebRequest)WebRequest.Create(url);
                req.Method = "GET";
                req.Timeout = Math.Max(1, timeoutSeconds) * 1000;
                req.ReadWriteTimeout = req.Timeout;
                using (var resp = (HttpWebResponse)req.GetResponse())
                {
                    if (resp.StatusCode != HttpStatusCode.OK) return false;
                    var ct = resp.ContentType ?? string.Empty;
                    if (!ct.StartsWith("image/", StringComparison.OrdinalIgnoreCase)) return false;
                    long len = resp.ContentLength;
                    if (len > 0 && len > maxBytes) return false;
                    using var s = resp.GetResponseStream();
                    if (s == null) return false;
                    using var ms = new MemoryStream(); s.CopyTo(ms);
                    if (ms.Length > maxBytes) return false;
                    var arr = ms.ToArray();
                    Cache[url] = arr;
                    bytes = arr; contentType = ct;
                    return true;
                }
#else
                using (var http = new HttpClient() { Timeout = TimeSpan.FromSeconds(Math.Max(1, timeoutSeconds)) })
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
