using OfficeIMO.Email.Store;

namespace OfficeIMO.Reader.EmailStore;

/// <summary>Options for projecting email stores through OfficeIMO.Reader.</summary>
public sealed class ReaderEmailStoreOptions {
    /// <summary>
    /// Gets or sets bounded store-reader options. Registrations capture a defensive copy.
    /// A Reader-level <see cref="ReaderOptions.MaxInputBytes"/> can narrow, but never widen, this limit.
    /// </summary>
    /// <remarks>
    /// A configured PST password is retained by the handler registration in memory so subsequent reads can use it.
    /// It is never added to Reader results, diagnostics, or metadata.
    /// </remarks>
    public EmailStoreReaderOptions? StoreOptions { get; set; }
}
