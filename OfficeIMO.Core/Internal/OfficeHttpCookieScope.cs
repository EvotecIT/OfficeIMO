using System;
using System.Net;
using System.Net.Http;
#if NET8_0_OR_GREATER
using System.Runtime.Versioning;
#elif NETSTANDARD2_0
using System.Runtime.InteropServices;
#endif
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Core.Internal;

/// <summary>Owns HTTP cookies for one operation that uses a shared transport.</summary>
internal sealed class OfficeHttpCookieScope {
    private CookieContainer? _cookies;

#if NET8_0_OR_GREATER
    [SupportedOSPlatformGuard("browser")]
    private static bool IsBrowser => OperatingSystem.IsBrowser();
#elif NETSTANDARD2_0
    private static bool IsBrowser => RuntimeInformation.IsOSPlatform(OSPlatform.Create("BROWSER"));
#else
    private static bool IsBrowser => false;
#endif

    /// <summary>Prevents a pooled native handler from retaining cookies across operations.</summary>
    internal static void DisableAutomaticCookies(HttpClientHandler handler) {
        // Browser Fetch manages its own credentials and does not support UseCookies.
        if (!IsBrowser) {
            handler.UseCookies = false;
        }
    }

    /// <summary>Sends the request and retains its response cookies for this operation.</summary>
    /// <remarks>The validated logical URI identifies the cookie origin when the request uses a pinned IP address.</remarks>
    internal async Task<HttpResponseMessage> SendAsync(
        HttpClient client,
        HttpRequestMessage request,
        Uri logicalUri,
        CancellationToken cancellationToken) {
        string? cookieHeader = Volatile.Read(ref _cookies)?.GetCookieHeader(logicalUri);
        if (!string.IsNullOrEmpty(cookieHeader)) {
            request.Headers.TryAddWithoutValidation("Cookie", cookieHeader);
        }

        HttpResponseMessage response = await client.SendAsync(
            request, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false);
        try {
            if (!IsBrowser) {
                StoreResponseCookies(response, logicalUri, "Set-Cookie");
#if NETFRAMEWORK
                StoreResponseCookies(response, logicalUri, "Set-Cookie2");
#endif
            }
            return response;
        } catch {
            response.Dispose();
            throw;
        }
    }

    private void StoreResponseCookies(HttpResponseMessage response, Uri logicalUri, string headerName) {
        if (!response.Headers.TryGetValues(headerName, out var headers)) {
            return;
        }

        CookieContainer cookies = LazyInitializer.EnsureInitialized(ref _cookies, static () => new CookieContainer())!;
        foreach (string header in headers) {
            try {
                cookies.SetCookies(logicalUri, header);
            } catch (CookieException) {
                // A malformed cookie must not prevent the resource from loading.
            }
        }
    }
}
