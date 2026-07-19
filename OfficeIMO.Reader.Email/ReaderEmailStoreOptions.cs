using OfficeIMO.Email.Store;

namespace OfficeIMO.Reader.Email;

/// <summary>Options for projecting email stores through OfficeIMO.Reader.</summary>
public sealed class ReaderEmailStoreOptions {
    private int _maxItems = 1_000;

    /// <summary>
    /// Gets or sets bounded store-reader options. Registrations capture a defensive copy.
    /// A Reader-level <see cref="ReaderOptions.MaxInputBytes"/> can narrow, but never widen, this limit.
    /// </summary>
    /// <remarks>
    /// A configured PST password is retained by the handler registration in memory so subsequent reads can use it.
    /// It is never added to Reader results, diagnostics, or metadata.
    /// </remarks>
    public EmailStoreReaderOptions? StoreOptions { get; set; }

    /// <summary>
    /// Optional selective summary query applied before full item projection. This is the preferred way to ingest
    /// a narrow slice of a very large PST or OST.
    /// </summary>
    public EmailStoreQuery? Query { get; set; }

    /// <summary>
    /// Selective item parts projected before Reader chunking. Null uses the full semantic item profile.
    /// </summary>
    public EmailStoreItemReadOptions? ItemReadOptions { get; set; }

    /// <summary>
    /// Whether PST/OST attachment payloads use session-bound streams instead of resident byte arrays.
    /// Enabled by default for bounded large-store ingestion.
    /// </summary>
    public bool StreamAttachmentContent { get; set; } = true;

    /// <summary>Maximum matching items fully projected into one Reader result. Default: 1,000.</summary>
    public int MaxItems {
        get => _maxItems;
        set {
            if (value <= 0) throw new ArgumentOutOfRangeException(nameof(value));
            _maxItems = value;
        }
    }

    /// <summary>Whether a corrupt or over-limit item is diagnosed and skipped instead of aborting the store.</summary>
    public bool ContinueOnItemError { get; set; } = true;

    /// <summary>
    /// Whether Reader hashes the complete source store. Disabled by default because hashing a huge PST/OST forces a
    /// full sequential read even when parsing selected items is random-access and bounded. Chunk hashes still follow
    /// <see cref="ReaderOptions.ComputeHashes"/>.
    /// </summary>
    public bool ComputeSourceHash { get; set; }
}
