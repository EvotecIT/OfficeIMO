using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    internal static class ExcelHttpWorkbookLoader {
        private const int BufferSize = 81920;
        private const int MaxRedirects = 10;

        internal static async Task<byte[]> DownloadAsync(
            Uri uri,
            ExcelHttpLoadOptions? options,
            CancellationToken cancellationToken = default,
            long? maximumBytes = null) {
            if (uri == null) throw new ArgumentNullException(nameof(uri));

            var snapshot = ExcelHttpLoadOptionsSnapshot.Create(options, maximumBytes);
            ValidateScheme(uri, snapshot.SchemePolicy);
            ValidateHost(uri, snapshot);
            ValidateLimits(snapshot);

            using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeoutCts.CancelAfter(snapshot.Timeout);

            using var client = CreateOwnedHttpClient(snapshot.Timeout, snapshot.HttpMessageHandler);
            using var response = await SendWithRedirectsAsync(client, uri, snapshot, timeoutCts.Token).ConfigureAwait(false);

            response.EnsureSuccessStatusCode();
            ValidateContentType(response, snapshot);

            long? contentLength = response.Content.Headers.ContentLength;
            if (contentLength.HasValue && contentLength.Value > snapshot.MaxBytes) {
                throw new IOException($"Remote workbook is too large. Content-Length is {contentLength.Value} bytes and the configured limit is {snapshot.MaxBytes} bytes.");
            }

            byte[] bytes = await ReadResponseBytesAsync(
                response,
                snapshot,
                contentLength,
                timeoutCts.Token).ConfigureAwait(false);

            if (snapshot.ValidateZipHeader && !LooksLikeZipPackage(bytes)) {
                throw new InvalidDataException("Downloaded workbook does not look like an Office Open XML package.");
            }

            return bytes;
        }

        private static HttpClient CreateOwnedHttpClient(TimeSpan timeout, HttpMessageHandler? messageHandler) {
            bool disposeHandler = messageHandler == null;
            messageHandler ??= new HttpClientHandler {
                AllowAutoRedirect = false,
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            };

            return new HttpClient(messageHandler, disposeHandler) {
                Timeout = timeout
            };
        }

        private static async Task<HttpResponseMessage> SendWithRedirectsAsync(
            HttpClient client,
            Uri initialUri,
            ExcelHttpLoadOptionsSnapshot options,
            CancellationToken cancellationToken) {
            Uri currentUri = initialUri;
            bool includeCustomHeaders = true;

            for (int redirectCount = 0; redirectCount <= MaxRedirects; redirectCount++) {
                var request = new HttpRequestMessage(HttpMethod.Get, currentUri);
                ApplyHeaders(request, options, includeCustomHeaders);

                HttpResponseMessage response;
                try {
                    response = await client.SendAsync(
                        request,
                        HttpCompletionOption.ResponseHeadersRead,
                        cancellationToken).ConfigureAwait(false);
                } finally {
                    request.Dispose();
                }

                if (!IsRedirect(response.StatusCode)) {
                    return response;
                }

                Uri nextUri = ResolveRedirectUri(currentUri, response.Headers.Location);
                response.Dispose();

                if (redirectCount == MaxRedirects) {
                    throw new HttpRequestException($"Remote workbook request exceeded the maximum of {MaxRedirects} redirects.");
                }

                ValidateScheme(nextUri, options.SchemePolicy);
                ValidateHost(nextUri, options);
                if (!IsSameOrigin(currentUri, nextUri)) {
                    includeCustomHeaders = false;
                }

                currentUri = nextUri;
            }

            throw new HttpRequestException($"Remote workbook request exceeded the maximum of {MaxRedirects} redirects.");
        }

        private static async Task<byte[]> ReadResponseBytesAsync(
            HttpResponseMessage response,
            ExcelHttpLoadOptionsSnapshot options,
            long? contentLength,
            CancellationToken cancellationToken) {
            using var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false);
            using var buffer = contentLength.HasValue && contentLength.Value <= int.MaxValue
                ? new MemoryStream((int)contentLength.Value)
                : new MemoryStream();

            byte[] chunk = new byte[BufferSize];
            long total = 0;

            while (true) {
                int read = await stream.ReadAsync(chunk, 0, chunk.Length, cancellationToken).ConfigureAwait(false);
                if (read == 0) {
                    break;
                }

                total += read;
                if (total > options.MaxBytes) {
                    throw new IOException($"Remote workbook exceeded the configured limit of {options.MaxBytes} bytes.");
                }

                buffer.Write(chunk, 0, read);
                options.Progress?.Report(new ExcelHttpLoadProgress(total, contentLength));
            }

            return buffer.ToArray();
        }

        private static void ValidateScheme(Uri uri, ExcelUriSchemePolicy schemePolicy) {
            if (string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)) {
                return;
            }

            if (schemePolicy == ExcelUriSchemePolicy.HttpAndHttps
                && string.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase)) {
                return;
            }

            throw new NotSupportedException("Remote Excel loads require HTTPS unless ExcelHttpLoadOptions.SchemePolicy allows HTTP.");
        }

        private static void ValidateHost(Uri uri, ExcelHttpLoadOptionsSnapshot options) {
            if (options.AllowedHosts.Count == 0) {
                return;
            }

            string host = NormalizeUriHost(uri);
            if (options.AllowedHosts.Contains(host)) {
                return;
            }

            throw new NotSupportedException($"Remote Excel load host '{uri.Host}' is not allowed by ExcelHttpLoadOptions.AllowedHosts.");
        }

        private static void ValidateLimits(ExcelHttpLoadOptionsSnapshot options) {
            if (options.MaxBytes <= 0) {
                throw new ArgumentOutOfRangeException(nameof(ExcelHttpLoadOptions.MaxBytes), "MaxBytes must be greater than zero.");
            }

            if (options.Timeout <= TimeSpan.Zero) {
                throw new ArgumentOutOfRangeException(nameof(ExcelHttpLoadOptions.Timeout), "Timeout must be greater than zero.");
            }
        }

        private static void ApplyHeaders(HttpRequestMessage request, ExcelHttpLoadOptionsSnapshot options, bool includeCustomHeaders) {
            if (includeCustomHeaders) {
                foreach (var header in options.Headers) {
                    if (!request.Headers.TryAddWithoutValidation(header.Key, header.Value)) {
                        throw new ArgumentException($"Header '{header.Key}' is not valid for an HTTP workbook request.");
                    }
                }
            }

            if (!string.IsNullOrWhiteSpace(options.UserAgent)
                && !options.Headers.ContainsKey("User-Agent")
                && !request.Headers.TryAddWithoutValidation("User-Agent", options.UserAgent)) {
                throw new ArgumentException("UserAgent is not valid for an HTTP workbook request.");
            }
        }

        private static bool IsRedirect(HttpStatusCode statusCode) {
            return statusCode == HttpStatusCode.Moved
                || statusCode == HttpStatusCode.Redirect
                || statusCode == HttpStatusCode.SeeOther
                || statusCode == HttpStatusCode.TemporaryRedirect
                || (int)statusCode == 308;
        }

        private static Uri ResolveRedirectUri(Uri currentUri, Uri? location) {
            if (location == null) {
                throw new HttpRequestException("Remote workbook redirect response did not include a Location header.");
            }

            return location.IsAbsoluteUri ? location : new Uri(currentUri, location);
        }

        private static bool IsSameOrigin(Uri left, Uri right) {
            return string.Equals(left.Scheme, right.Scheme, StringComparison.OrdinalIgnoreCase)
                && string.Equals(left.IdnHost, right.IdnHost, StringComparison.OrdinalIgnoreCase)
                && left.Port == right.Port;
        }

        private static void ValidateContentType(HttpResponseMessage response, ExcelHttpLoadOptionsSnapshot options) {
            if (!options.ValidateContentTypeWhenPresent) {
                return;
            }

            string? mediaType = response.Content.Headers.ContentType?.MediaType;
            if (string.IsNullOrWhiteSpace(mediaType)) {
                return;
            }

            if (!options.AllowedContentTypes.Contains(mediaType!)) {
                throw new InvalidDataException($"Remote workbook response Content-Type '{mediaType}' is not allowed.");
            }
        }

        private static bool LooksLikeZipPackage(byte[] bytes) {
            return bytes.Length >= 2 && bytes[0] == (byte)'P' && bytes[1] == (byte)'K';
        }

        private static HashSet<string> NormalizeAllowedHosts(ISet<string> hosts) {
            var normalizedHosts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string host in hosts) {
                normalizedHosts.Add(NormalizeConfiguredHost(host));
            }

            return normalizedHosts;
        }

        private static string NormalizeConfiguredHost(string host) {
            if (string.IsNullOrWhiteSpace(host)) {
                throw new ArgumentException("AllowedHosts entries must be non-empty host names.", nameof(ExcelHttpLoadOptions.AllowedHosts));
            }

            string candidate = host.Trim();
            if (candidate.Length > 1 && candidate[0] == '[' && candidate[candidate.Length - 1] == ']') {
                candidate = candidate.Substring(1, candidate.Length - 2);
            }

            UriHostNameType hostNameType = Uri.CheckHostName(candidate);
            if (hostNameType == UriHostNameType.Unknown) {
                throw new ArgumentException("AllowedHosts entries must be host names only; schemes, paths, and ports are not accepted.", nameof(ExcelHttpLoadOptions.AllowedHosts));
            }

            if (hostNameType == UriHostNameType.IPv6) {
                return candidate.ToLowerInvariant();
            }

            return new Uri("https://" + candidate).IdnHost.TrimEnd('.').ToLowerInvariant();
        }

        private static string NormalizeUriHost(Uri uri) {
            string host = uri.IdnHost;
            if (Uri.CheckHostName(host) == UriHostNameType.IPv6) {
                return host.Trim('[', ']').ToLowerInvariant();
            }

            return host.TrimEnd('.').ToLowerInvariant();
        }

        private sealed class ExcelHttpLoadOptionsSnapshot {
            private ExcelHttpLoadOptionsSnapshot(
                ExcelUriSchemePolicy schemePolicy,
                long maxBytes,
                TimeSpan timeout,
                string? userAgent,
                Dictionary<string, string> headers,
                HashSet<string> allowedHosts,
                bool validateZipHeader,
                bool validateContentTypeWhenPresent,
                HashSet<string> allowedContentTypes,
                IProgress<ExcelHttpLoadProgress>? progress,
                HttpMessageHandler? httpMessageHandler) {
                SchemePolicy = schemePolicy;
                MaxBytes = maxBytes;
                Timeout = timeout;
                UserAgent = userAgent;
                Headers = headers;
                AllowedHosts = allowedHosts;
                ValidateZipHeader = validateZipHeader;
                ValidateContentTypeWhenPresent = validateContentTypeWhenPresent;
                AllowedContentTypes = allowedContentTypes;
                Progress = progress;
                HttpMessageHandler = httpMessageHandler;
            }

            internal ExcelUriSchemePolicy SchemePolicy { get; }
            internal long MaxBytes { get; }
            internal TimeSpan Timeout { get; }
            internal string? UserAgent { get; }
            internal Dictionary<string, string> Headers { get; }
            internal HashSet<string> AllowedHosts { get; }
            internal bool ValidateZipHeader { get; }
            internal bool ValidateContentTypeWhenPresent { get; }
            internal HashSet<string> AllowedContentTypes { get; }
            internal IProgress<ExcelHttpLoadProgress>? Progress { get; }
            internal HttpMessageHandler? HttpMessageHandler { get; }

            internal static ExcelHttpLoadOptionsSnapshot Create(ExcelHttpLoadOptions? options, long? maximumBytes) {
                options ??= new ExcelHttpLoadOptions();
                long maxBytes = maximumBytes.HasValue
                    ? Math.Min(options.MaxBytes, maximumBytes.Value)
                    : options.MaxBytes;

                return new ExcelHttpLoadOptionsSnapshot(
                    options.SchemePolicy,
                    maxBytes,
                    options.Timeout,
                    options.UserAgent,
                    new Dictionary<string, string>(options.Headers, StringComparer.OrdinalIgnoreCase),
                    NormalizeAllowedHosts(options.AllowedHosts),
                    options.ValidateZipHeader,
                    options.ValidateContentTypeWhenPresent,
                    new HashSet<string>(options.AllowedContentTypes, StringComparer.OrdinalIgnoreCase),
                    options.Progress,
                    options.HttpMessageHandler);
            }
        }
    }
}
