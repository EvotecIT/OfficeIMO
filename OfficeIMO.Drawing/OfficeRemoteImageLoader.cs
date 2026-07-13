using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Drawing;

/// <summary>
/// Retrieves bounded image content for OfficeIMO packages that explicitly opt into remote I/O.
/// </summary>
public static class OfficeRemoteImageLoader {
    private const int BufferSize = 81920;
    private static readonly HttpClient Client = CreateClient();

    /// <summary>
    /// Asynchronously retrieves an image from an HTTP or HTTPS URL.
    /// </summary>
    public static Task<OfficeRemoteImage> LoadAsync(
        string url,
        OfficeRemoteImageLoadOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(url)) {
            throw new ArgumentException("Image URL cannot be null or whitespace.", nameof(url));
        }

        if (!Uri.TryCreate(url, UriKind.Absolute, out Uri? uri)) {
            throw new ArgumentException("Image URL must be an absolute HTTP or HTTPS URL.", nameof(url));
        }

        return LoadAsync(uri, options, cancellationToken);
    }

    /// <summary>
    /// Asynchronously retrieves an image from an HTTP or HTTPS URI.
    /// </summary>
    public static async Task<OfficeRemoteImage> LoadAsync(
        Uri uri,
        OfficeRemoteImageLoadOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (uri == null) throw new ArgumentNullException(nameof(uri));
        ValidateHttpUri(uri, nameof(uri));

        OfficeRemoteImageLoadOptions.Snapshot snapshot =
            (options ?? new OfficeRemoteImageLoadOptions()).CreateSnapshot();
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        if (snapshot.Timeout != System.Threading.Timeout.InfiniteTimeSpan) {
            timeout.CancelAfter(snapshot.Timeout);
        }

        Uri current = uri;
        for (int redirectCount = 0; ; redirectCount++) {
            using HttpResponseMessage response = await Client.GetAsync(
                current,
                HttpCompletionOption.ResponseHeadersRead,
                timeout.Token).ConfigureAwait(false);

            if (IsRedirect(response.StatusCode)) {
                if (redirectCount >= snapshot.MaximumRedirects) {
                    throw new InvalidDataException($"The remote image exceeded the {snapshot.MaximumRedirects}-redirect limit.");
                }

                Uri next = ResolveRedirect(current, response.Headers.Location);
                current = next;
                continue;
            }

            response.EnsureSuccessStatusCode();
            string contentType = NormalizeImageContentType(response.Content.Headers.ContentType?.MediaType);
            long? contentLength = response.Content.Headers.ContentLength;
            if (contentLength.HasValue && contentLength.Value > snapshot.MaximumBytes) {
                throw new InvalidDataException($"The remote image exceeds the {snapshot.MaximumBytes}-byte limit.");
            }

            using Stream input = await response.Content.ReadAsStreamAsync().ConfigureAwait(false);
            byte[] bytes = await ReadBoundedAsync(input, snapshot.MaximumBytes, timeout.Token).ConfigureAwait(false);
            string fileName = Path.GetFileName(current.LocalPath);
            if (string.IsNullOrWhiteSpace(fileName)) {
                fileName = "image" + OfficeImageInfo.GetDefaultExtension(OfficeImageInfo.FromMimeType(contentType));
            }

            return new OfficeRemoteImage(current, bytes, fileName, contentType);
        }
    }

    private static HttpClient CreateClient() {
        var handler = new HttpClientHandler {
            AllowAutoRedirect = false,
            AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
        };
        return new HttpClient(handler, disposeHandler: true) {
            Timeout = System.Threading.Timeout.InfiniteTimeSpan
        };
    }

    private static async Task<byte[]> ReadBoundedAsync(Stream input, long maximumBytes, CancellationToken cancellationToken) {
        using var output = new MemoryStream();
        var buffer = new byte[BufferSize];
        long total = 0;
        while (true) {
            int read = await input.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false);
            if (read == 0) break;

            total += read;
            if (total > maximumBytes) {
                throw new InvalidDataException($"The remote image exceeds the {maximumBytes}-byte limit.");
            }

            await output.WriteAsync(buffer, 0, read, cancellationToken).ConfigureAwait(false);
        }

        return output.ToArray();
    }

    private static string NormalizeImageContentType(string? contentType) {
        if (!OfficeImageInfo.TryNormalizeImageContentType(contentType, out string normalized)) {
            throw new InvalidDataException("The remote resource did not return an image content type.");
        }

        return normalized;
    }

    private static Uri ResolveRedirect(Uri current, Uri? location) {
        if (location == null) {
            throw new InvalidDataException("The remote image response contained a redirect without a location.");
        }

        Uri resolved = location.IsAbsoluteUri ? location : new Uri(current, location);
        ValidateHttpUri(resolved, nameof(location));
        if (!IsSameOrigin(current, resolved)) {
            throw new InvalidDataException("Remote image redirects must remain on the same origin.");
        }

        return resolved;
    }

    private static void ValidateHttpUri(Uri uri, string parameterName) {
        if (!uri.IsAbsoluteUri
            || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps)) {
            throw new ArgumentException("Image URI must use HTTP or HTTPS.", parameterName);
        }
    }

    private static bool IsSameOrigin(Uri left, Uri right) =>
        string.Equals(left.Scheme, right.Scheme, StringComparison.OrdinalIgnoreCase)
        && string.Equals(left.IdnHost, right.IdnHost, StringComparison.OrdinalIgnoreCase)
        && left.Port == right.Port;

    private static bool IsRedirect(HttpStatusCode statusCode) =>
        statusCode == HttpStatusCode.Moved
        || statusCode == HttpStatusCode.Redirect
        || statusCode == HttpStatusCode.SeeOther
        || statusCode == HttpStatusCode.TemporaryRedirect
        || (int)statusCode == 308;
}
