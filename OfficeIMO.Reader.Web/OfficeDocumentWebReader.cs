namespace OfficeIMO.Reader.Web;

/// <summary>
/// Bounded, thread-safe HTTP transport over a caller-configured <see cref="OfficeDocumentReader"/>.
/// </summary>
/// <remarks>
/// The caller owns the Reader and <see cref="HttpClient"/> lifetimes. This type does not mutate or dispose either.
/// URI validation provides defense in depth for schemes, allowlists, credentials, and non-public IP literals;
/// it is not an SSRF isolation boundary. When reading an untrusted URI, configure the caller-owned HTTP handler
/// to validate resolved addresses at connection time and validate each redirect before sending it.
/// </remarks>
public sealed class OfficeDocumentWebReader {
    private readonly OfficeDocumentReader _reader;
    private readonly HttpClient _httpClient;
    private readonly ReaderWebOptions _options;
    private readonly SemaphoreSlim _requestGate;

    /// <summary>Creates a bounded web reader over existing Reader and HTTP client instances.</summary>
    public OfficeDocumentWebReader(
        OfficeDocumentReader reader,
        HttpClient httpClient,
        ReaderWebOptions? options = null) {
        _reader = reader ?? throw new ArgumentNullException(nameof(reader));
        _httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
        _options = (options ?? new ReaderWebOptions()).CloneValidated();
        _requestGate = new SemaphoreSlim(_options.MaxConcurrentRequests, _options.MaxConcurrentRequests);
    }

    /// <summary>
    /// Downloads one bounded HTTP(S) response and routes its body through the configured Reader instance.
    /// </summary>
    /// <remarks>
    /// The URI must be trusted unless the caller-owned HTTP handler enforces connection-time address and
    /// pre-request redirect policy. Reader Web cannot inspect or override those behaviors on an existing client.
    /// </remarks>
    public async Task<OfficeDocumentReadResult> ReadDocumentAsync(
        Uri uri,
        string? sourceName = null,
        ReaderOptions? readerOptions = null,
        CancellationToken cancellationToken = default) {
        if (uri == null) throw new ArgumentNullException(nameof(uri));
        ReaderWebUriPolicy.Validate(uri, _options);
        string? normalizedSourceName = ReaderWebTransport.NormalizeExplicitSourceName(sourceName);
        await _requestGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            using ReaderWebDownload download = await ReaderWebTransport.DownloadAsync(
                _httpClient,
                uri,
                normalizedSourceName,
                logicalSourceName => GetEffectiveMaxInputBytes(logicalSourceName, readerOptions),
                _options,
                cancellationToken).ConfigureAwait(false);
            OfficeDocumentReadResult result = await _reader.ReadDocumentAsync(
                download.Content,
                download.SourceName,
                readerOptions,
                cancellationToken).ConfigureAwait(false);
            download.ApplyTransportMetadata(
                result,
                _options,
                readerOptions?.ComputeHashes ?? true);
            return result;
        } finally {
            _requestGate.Release();
        }
    }

    private long GetEffectiveMaxInputBytes(string sourceName, ReaderOptions? readerOptions) {
        long maxInputBytes = _options.MaxResponseBytes;
        if (readerOptions?.MaxInputBytes is long readerLimit && readerLimit >= 0) {
            maxInputBytes = Math.Min(maxInputBytes, readerLimit);
        }
        if (_reader.GetHandlerDefaultMaxInputBytes(sourceName) is long handlerLimit) {
            maxInputBytes = Math.Min(maxInputBytes, handlerLimit);
        }
        return maxInputBytes;
    }

    /// <summary>
    /// Downloads one bounded HTTP(S) response and returns the Markdown emitted by the same rich Reader pipeline.
    /// </summary>
    public async Task<string> ConvertToMarkdownAsync(
        Uri uri,
        string? sourceName = null,
        ReaderOptions? readerOptions = null,
        CancellationToken cancellationToken = default) {
        OfficeDocumentReadResult result = await ReadDocumentAsync(
            uri,
            sourceName,
            readerOptions,
            cancellationToken).ConfigureAwait(false);
        return result.Markdown ?? string.Empty;
    }
}
