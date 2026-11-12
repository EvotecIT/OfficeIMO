using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;

namespace OfficeIMO.Workflows;

/// <summary>Carries IPP messages over HTTP POST (RFC 8010 section 4).</summary>
internal sealed class IppHttpTransport : IDisposable {
    private const string IppMediaType = "application/ipp";
    private readonly HttpClient _client;

    public IppHttpTransport(Uri printerUri, HttpMessageHandler handler, bool disposeHandler) {
        HttpUri = ToHttpUri(printerUri);
        _client = new HttpClient(handler, disposeHandler) { Timeout = Timeout.InfiniteTimeSpan };
    }

    public Uri HttpUri { get; }

    /// <summary>Maps ipp:// to http:// and ipps:// to https://, using port 631 when the URI has none.</summary>
    public static Uri ToHttpUri(Uri printerUri) {
        string scheme = printerUri.Scheme.Equals("ipps", StringComparison.OrdinalIgnoreCase) ? "https" : "http";
        int port = printerUri.Port > 0 ? printerUri.Port : 631;
        return new UriBuilder(scheme, printerUri.Host, port, printerUri.AbsolutePath) { Query = printerUri.Query.TrimStart('?') }.Uri;
    }

    public static HttpMessageHandler CreateDefaultHandler() => new HttpClientHandler {
        AllowAutoRedirect = false,
        UseCookies = false,
        UseProxy = false,
        AutomaticDecompression = DecompressionMethods.None
    };

    /// <summary>Sends one request and returns the decoded response.</summary>
    /// <param name="request">Encoded IPP request header and attributes.</param>
    /// <param name="document">Document data that follows the attributes, streamed from its current position.</param>
    /// <param name="bodyStarted">Invoked when the request body begins to reach the connection.</param>
    /// <param name="timeout">Limit for the whole exchange.</param>
    /// <param name="maxResponseBytes">Largest accepted response.</param>
    /// <param name="cancellationToken">Caller cancellation.</param>
    /// <exception cref="IppException">The printer could not be reached, refused at the HTTP level, timed out, or sent a malformed response.</exception>
    public async Task<IppMessage> SendAsync(IppMessage request, Stream? document, Action? bodyStarted, TimeSpan timeout,
        int maxResponseBytes, CancellationToken cancellationToken) {
        byte[] header = IppCodec.Encode(request);
        using var timeoutSource = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        if (timeout != Timeout.InfiniteTimeSpan) timeoutSource.CancelAfter(timeout);
        CancellationToken token = timeoutSource.Token;
        try {
            using var content = new IppContent(header, document, bodyStarted, token);
            using var message = new HttpRequestMessage(HttpMethod.Post, HttpUri) { Content = content };
            message.Headers.ExpectContinue = false;
            if (!content.HasKnownLength) message.Headers.TransferEncodingChunked = true;
            using HttpResponseMessage response = await _client.SendAsync(message, HttpCompletionOption.ResponseHeadersRead, token).ConfigureAwait(false);
            int status = (int)response.StatusCode;
            if (status is >= 300 and < 400) {
                throw new IppException($"The printer at {HttpUri} redirected the request to '{response.Headers.Location}'. Redirects are refused; use the printer's own URI.", definitive: true);
            }
            if (response.RequestMessage?.RequestUri is { } answered && answered != HttpUri) {
                // A caller-supplied handler followed a redirect; the answer is not from the requested printer.
                throw new IppException($"The printer at {HttpUri} redirected the request to '{answered}'. Redirects are refused; use the printer's own URI.", definitive: false);
            }
            if (status == 401 || status == 407) {
                throw new IppException($"The printer at {HttpUri} requires authentication, which this client does not provide.", definitive: true);
            }
            if (status is < 200 or >= 300) {
                throw new IppException($"The printer at {HttpUri} returned HTTP {status} {response.ReasonPhrase}.", definitive: true);
            }
            byte[] body = await ReadBoundedAsync(response, maxResponseBytes, token).ConfigureAwait(false);
            IppMessage reply = IppCodec.Decode(body, body.Length, new IppDecodeLimits { MaxMessageBytes = maxResponseBytes }, out _);
            if (reply.RequestId != request.RequestId) throw IppCodec.Malformed($"response request-id {reply.RequestId} does not match {request.RequestId}.");
            return reply;
        } catch (Exception error) when (error is not IppException && IsTransportFailure(error)) {
            if (cancellationToken.IsCancellationRequested) throw new OperationCanceledException(cancellationToken);
            if (timeoutSource.IsCancellationRequested) {
                throw new IppException($"The printer at {HttpUri} did not complete the request within {timeout}.", null,
                    new TimeoutException("The IPP request timed out.", error));
            }
            throw new IppException($"Could not communicate with the printer at {HttpUri}: {error.Message}", null, error);
        }
    }

    private static bool IsTransportFailure(Exception error) =>
        error is HttpRequestException or IOException or OperationCanceledException or ObjectDisposedException;

    private static async Task<byte[]> ReadBoundedAsync(HttpResponseMessage response, int maxBytes, CancellationToken token) {
        if (response.Content.Headers.ContentLength is long declared && declared > maxBytes) {
            throw new IppException($"The printer response declares {declared} bytes, more than the {maxBytes}-byte limit.");
        }
        // Stream reads on older frameworks may ignore the token, so cancellation also closes the response.
        using CancellationTokenRegistration registration = token.Register(response.Dispose);
        using Stream stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false);
        using var buffer = new MemoryStream();
        var chunk = new byte[16384];
        while (true) {
            int read = await stream.ReadAsync(chunk, 0, chunk.Length, token).ConfigureAwait(false);
            if (read == 0) return buffer.ToArray();
            if (buffer.Length + read > maxBytes) throw new IppException($"The printer response exceeds the {maxBytes}-byte limit.");
            buffer.Write(chunk, 0, read);
        }
    }

    public void Dispose() => _client.Dispose();

    /// <summary>The request body: encoded attributes followed by streamed document data.</summary>
    private sealed class IppContent : HttpContent {
        private readonly byte[] _header;
        private readonly Stream? _document;
        private readonly Action? _started;
        private readonly CancellationToken _token;
        private readonly long? _length;

        public IppContent(byte[] header, Stream? document, Action? started, CancellationToken token) {
            _header = header;
            _document = document;
            _started = started;
            _token = token;
            if (document is null) _length = header.Length;
            else if (document.CanSeek) _length = header.Length + Math.Max(0, document.Length - document.Position);
            Headers.ContentType = new MediaTypeHeaderValue(IppMediaType);
        }

        public bool HasKnownLength => _length.HasValue;

        protected override Task SerializeToStreamAsync(Stream stream, TransportContext? context) => WriteAsync(stream, _token);

#if NET5_0_OR_GREATER
        protected override Task SerializeToStreamAsync(Stream stream, TransportContext? context, CancellationToken cancellationToken) =>
            WriteAsync(stream, cancellationToken.CanBeCanceled ? cancellationToken : _token);
#endif

        private async Task WriteAsync(Stream stream, CancellationToken token) {
            _started?.Invoke();
            await stream.WriteAsync(_header, 0, _header.Length, token).ConfigureAwait(false);
            if (_document is null) return;
            var buffer = new byte[81920];
            long remaining = _length.HasValue ? _length.Value - _header.Length : long.MaxValue;
            while (remaining > 0) {
                int read = await _document.ReadAsync(buffer, 0, (int)Math.Min(buffer.Length, remaining), token).ConfigureAwait(false);
                if (read == 0) break;
                await stream.WriteAsync(buffer, 0, read, token).ConfigureAwait(false);
                remaining -= read;
            }
            if (_length.HasValue && remaining > 0) throw new IOException("The document stream ended before its declared length.");
        }

        protected override bool TryComputeLength(out long length) {
            length = _length ?? 0;
            return _length.HasValue;
        }
    }
}
