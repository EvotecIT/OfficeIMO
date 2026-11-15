using OfficeIMO.Core.Internal;

namespace OfficeIMO.Reader.Web;

/// <summary>
/// Bounded, thread-safe HTTP transport over a caller-configured <see cref="OfficeDocumentReader"/>.
/// </summary>
/// <remarks>
/// The caller owns the Reader lifetime. A web reader created without an <see cref="HttpClient"/> uses a shared
/// transport that refuses loopback, private, link-local and other non-public addresses after name resolution and
/// on every redirect, unless <see cref="ReaderWebOptions.AllowPrivateNetworkAddresses"/> is set. A caller-supplied
/// client stays caller-owned; its handler resolves names, connects and may follow redirects itself, so Reader Web
/// screens only IP literals and localhost names for it.
/// </remarks>
public sealed class OfficeDocumentWebReader {
    private readonly OfficeDocumentReader _reader;
    private readonly HttpClient _httpClient;
    private readonly bool _ownedTransport;
    private readonly ReaderWebOptions _options;
    private readonly SemaphoreSlim _requestGate;

    /// <summary>
    /// Creates a bounded web reader over an existing Reader that downloads through OfficeIMO's address-checked transport.
    /// </summary>
    public OfficeDocumentWebReader(
        OfficeDocumentReader reader,
        ReaderWebOptions? options = null) {
        _reader = reader ?? throw new ArgumentNullException(nameof(reader));
        _options = (options ?? new ReaderWebOptions()).CloneValidated();
        _httpClient = OfficeNetworkAddressPolicy.GetClient(_options.AllowPrivateNetworkAddresses);
        _ownedTransport = true;
        _requestGate = new SemaphoreSlim(_options.MaxConcurrentRequests, _options.MaxConcurrentRequests);
    }

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
    /// Reader Web follows and checks redirects that reach it. A caller-supplied handler that follows redirects
    /// itself sends the redirected request before Reader Web sees it.
    /// </remarks>
    public async Task<OfficeDocumentReadResult> ReadDocumentAsync(
        Uri uri,
        string? sourceName = null,
        ReaderOptions? readerOptions = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(uri);
        ReaderWebUriPolicy.Validate(uri, _options, screenHostNames: !_ownedTransport);
        string? normalizedSourceName = ReaderWebTransport.NormalizeExplicitSourceName(sourceName);
        await _requestGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            using ReaderWebDownload download = await ReaderWebTransport.DownloadAsync(
                _httpClient,
                _ownedTransport,
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
