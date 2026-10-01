using System.Net.Http.Headers;
using System.Text.Json;

namespace OfficeIMO.Invoicing.KSeF;

/// <summary>A bounded KSeF v2 client. It does not retry mutations or follow redirects. An injected transport is a trusted boundary and must preserve these rules.</summary>
public sealed partial class KsefClient : IDisposable {
    private readonly HttpClient _http;
    private readonly Uri _baseUri;
    private readonly Guid _owner = Guid.NewGuid();
    private readonly TimeProvider _clock;
    private readonly TimeSpan _timeout;
    /// <summary>Creates a client for an explicit official environment. The optional handler and clock support integration hosts and credential-free tests.</summary>
    public KsefClient(KsefEnvironment environment = KsefEnvironment.Test, HttpMessageHandler? handler = null, TimeProvider? clock = null, TimeSpan? requestTimeout = null) {
        _baseUri = new Uri(environment switch { KsefEnvironment.Test => "https://api-test.ksef.mf.gov.pl/v2/", KsefEnvironment.Demo => "https://api-demo.ksef.mf.gov.pl/v2/", KsefEnvironment.Production => "https://api.ksef.mf.gov.pl/v2/", _ => throw new ArgumentOutOfRangeException(nameof(environment)) });
        Environment = environment; _clock = clock ?? TimeProvider.System; _timeout = requestTimeout ?? TimeSpan.FromSeconds(30);
        if (_timeout < TimeSpan.FromSeconds(1) || _timeout > TimeSpan.FromMinutes(5)) throw new ArgumentOutOfRangeException(nameof(requestTimeout));
        _http = new HttpClient(handler ?? new SocketsHttpHandler { AllowAutoRedirect = false, MaxResponseHeadersLength = 32 }, true) { Timeout = Timeout.InfiniteTimeSpan };
    }
    /// <summary>Selected official environment.</summary>
    public KsefEnvironment Environment { get; }
    private void Owned(Guid owner) { if (owner != _owner) throw new InvalidOperationException("Handle belongs to a different KSeF client; contexts and environments cannot be mixed."); }
    private string Authorization(KsefCredentials credentials, bool refresh = false) { ArgumentNullException.ThrowIfNull(credentials); Owned(credentials.Owner); return credentials.Header(_clock.GetUtcNow(), refresh); }
    private static bool ProtocolFailure(Exception error) => error is InvalidDataException or JsonException or KeyNotFoundException or InvalidOperationException or ArgumentException or FormatException or OverflowException;
    private async Task<T> JsonAsync<T>(HttpMethod method, string path, byte[]? payload, string? token, Func<JsonElement, T> parse, CancellationToken cancellationToken, bool mutation = false, string? reference = null, string? invoiceHash = null, Action? beforeDispatch = null, string? continuation = null) {
        byte[] bytes = await SendAsync(method, path, payload, token, 1024 * 1024, cancellationToken, mutation, reference, invoiceHash, beforeDispatch, continuation).ConfigureAwait(false);
        try { using JsonDocument document = JsonDocument.Parse(bytes, new JsonDocumentOptions { MaxDepth = 32 }); return parse(document.RootElement); }
        catch (Exception error) when (ProtocolFailure(error)) { if (mutation) throw new KsefMutationAmbiguousException(path, reference, invoiceHash); throw new InvalidDataException("KSeF returned an invalid or unsupported response."); }
        finally { System.Security.Cryptography.CryptographicOperations.ZeroMemory(bytes); }
    }
    private async Task<byte[]> SendAsync(HttpMethod method, string path, byte[]? payload, string? token, int maximum, CancellationToken cancellationToken, bool mutation = false, string? reference = null, string? invoiceHash = null, Action? beforeDispatch = null, string? continuation = null) {
        cancellationToken.ThrowIfCancellationRequested();
        using var request = new HttpRequestMessage(method, new Uri(_baseUri, path));
        if (token != null) request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        if (continuation != null) {
            if (continuation.Length > 8192 || continuation.Any(char.IsControl)) throw new ArgumentException("Continuation token exceeds the supported header bound.", nameof(continuation));
            request.Headers.Add("x-continuation-token", continuation);
        }
        if (payload != null) { request.Content = new ByteArrayContent(payload); request.Content.Headers.ContentType = new MediaTypeHeaderValue("application/json"); }
        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken); deadline.CancelAfter(_timeout);
        beforeDispatch?.Invoke(); bool dispatched = false;
        try {
            deadline.Token.ThrowIfCancellationRequested(); dispatched = true;
            using HttpResponseMessage response = await _http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, deadline.Token).ConfigureAwait(false);
            if (response.RequestMessage?.RequestUri != null && response.RequestMessage.RequestUri != request.RequestUri) throw new InvalidDataException("Transport redirected a KSeF request.");
            if (!response.IsSuccessStatusCode) {
                if (mutation && ((int)response.StatusCode >= 500 || response.StatusCode == System.Net.HttpStatusCode.RequestTimeout)) throw new KsefMutationAmbiguousException(path, reference, invoiceHash);
                TimeSpan? retry = response.Headers.RetryAfter?.Delta;
                if (retry == null && response.Headers.RetryAfter?.Date is DateTimeOffset retryAt) retry = retryAt > _clock.GetUtcNow() ? retryAt - _clock.GetUtcNow() : TimeSpan.Zero;
                throw new KsefApiException(response.StatusCode, retry);
            }
            if (response.Content.Headers.ContentLength > maximum) throw new InvalidDataException("KSeF response exceeds the supported size bound.");
            using Stream input = await response.Content.ReadAsStreamAsync(deadline.Token).ConfigureAwait(false);
            using var output = new MemoryStream(); byte[] buffer = new byte[8192];
            try {
                while (true) {
                    int read = await input.ReadAsync(buffer, deadline.Token).ConfigureAwait(false); if (read == 0) break;
                    if (output.Length + read > maximum) throw new InvalidDataException("KSeF response exceeds the supported size bound.");
                    output.Write(buffer, 0, read);
                }
                return output.ToArray();
            } finally {
                System.Security.Cryptography.CryptographicOperations.ZeroMemory(buffer);
                System.Security.Cryptography.CryptographicOperations.ZeroMemory(output.GetBuffer());
            }
        } catch (Exception error) when (error is HttpRequestException or IOException or InvalidDataException or OperationCanceledException) {
            if (mutation && dispatched) throw new KsefMutationAmbiguousException(path, reference, invoiceHash);
            throw;
        }
    }
    /// <summary>Disposes the owned HTTP transport; caller-owned credentials and sessions must be disposed separately.</summary>
    public void Dispose() => _http.Dispose();
}
