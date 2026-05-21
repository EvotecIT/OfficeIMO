using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    internal static class ExcelHttpWorkbookLoader {
        private const int BufferSize = 81920;

        internal static byte[] Download(Uri uri, ExcelHttpLoadOptions? options, CancellationToken cancellationToken = default) {
            return DownloadAsync(uri, options, cancellationToken).GetAwaiter().GetResult();
        }

        internal static async Task<byte[]> DownloadAsync(Uri uri, ExcelHttpLoadOptions? options, CancellationToken cancellationToken = default) {
            if (uri == null) throw new ArgumentNullException(nameof(uri));

            var snapshot = ExcelHttpLoadOptionsSnapshot.Create(options);
            ValidateScheme(uri, snapshot.SchemePolicy);
            ValidateLimits(snapshot);

            using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeoutCts.CancelAfter(snapshot.Timeout);

            HttpClient? ownedClient = null;
            var client = snapshot.HttpClient;
            if (client == null) {
                ownedClient = new HttpClient();
                client = ownedClient;
            }

            try {
                using var request = new HttpRequestMessage(HttpMethod.Get, uri);
                ApplyHeaders(request, snapshot);

                using var response = await client.SendAsync(
                    request,
                    HttpCompletionOption.ResponseHeadersRead,
                    timeoutCts.Token).ConfigureAwait(false);

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
            } finally {
                ownedClient?.Dispose();
            }
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

        private static void ValidateLimits(ExcelHttpLoadOptionsSnapshot options) {
            if (options.MaxBytes <= 0) {
                throw new ArgumentOutOfRangeException(nameof(ExcelHttpLoadOptions.MaxBytes), "MaxBytes must be greater than zero.");
            }

            if (options.Timeout <= TimeSpan.Zero) {
                throw new ArgumentOutOfRangeException(nameof(ExcelHttpLoadOptions.Timeout), "Timeout must be greater than zero.");
            }
        }

        private static void ApplyHeaders(HttpRequestMessage request, ExcelHttpLoadOptionsSnapshot options) {
            foreach (var header in options.Headers) {
                if (!request.Headers.TryAddWithoutValidation(header.Key, header.Value)) {
                    throw new ArgumentException($"Header '{header.Key}' is not valid for an HTTP workbook request.");
                }
            }

            if (!string.IsNullOrWhiteSpace(options.UserAgent)
                && !options.Headers.ContainsKey("User-Agent")
                && !request.Headers.TryAddWithoutValidation("User-Agent", options.UserAgent)) {
                throw new ArgumentException("UserAgent is not valid for an HTTP workbook request.");
            }
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

        private sealed class ExcelHttpLoadOptionsSnapshot {
            private ExcelHttpLoadOptionsSnapshot(
                ExcelUriSchemePolicy schemePolicy,
                long maxBytes,
                TimeSpan timeout,
                string? userAgent,
                Dictionary<string, string> headers,
                bool validateZipHeader,
                bool validateContentTypeWhenPresent,
                HashSet<string> allowedContentTypes,
                IProgress<ExcelHttpLoadProgress>? progress,
                HttpClient? httpClient) {
                SchemePolicy = schemePolicy;
                MaxBytes = maxBytes;
                Timeout = timeout;
                UserAgent = userAgent;
                Headers = headers;
                ValidateZipHeader = validateZipHeader;
                ValidateContentTypeWhenPresent = validateContentTypeWhenPresent;
                AllowedContentTypes = allowedContentTypes;
                Progress = progress;
                HttpClient = httpClient;
            }

            internal ExcelUriSchemePolicy SchemePolicy { get; }
            internal long MaxBytes { get; }
            internal TimeSpan Timeout { get; }
            internal string? UserAgent { get; }
            internal Dictionary<string, string> Headers { get; }
            internal bool ValidateZipHeader { get; }
            internal bool ValidateContentTypeWhenPresent { get; }
            internal HashSet<string> AllowedContentTypes { get; }
            internal IProgress<ExcelHttpLoadProgress>? Progress { get; }
            internal HttpClient? HttpClient { get; }

            internal static ExcelHttpLoadOptionsSnapshot Create(ExcelHttpLoadOptions? options) {
                options ??= new ExcelHttpLoadOptions();

                return new ExcelHttpLoadOptionsSnapshot(
                    options.SchemePolicy,
                    options.MaxBytes,
                    options.Timeout,
                    options.UserAgent,
                    new Dictionary<string, string>(options.Headers, StringComparer.OrdinalIgnoreCase),
                    options.ValidateZipHeader,
                    options.ValidateContentTypeWhenPresent,
                    new HashSet<string>(options.AllowedContentTypes, StringComparer.OrdinalIgnoreCase),
                    options.Progress,
                    options.HttpClient);
            }
        }
    }
}
