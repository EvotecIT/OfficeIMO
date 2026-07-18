namespace OfficeIMO.Email.Store;

public sealed partial class EmailStoreSession {
    /// <summary>
    /// Asynchronously streams lightweight references without adding an async-interface compatibility dependency.
    /// The returned concrete type supports C# <c>await foreach</c> on every package target.
    /// </summary>
    public EmailStoreAsyncEnumerable<EmailStoreItemReference> EnumerateItemsAsync(
        EmailStoreEnumerationOptions? options = null,
        CancellationToken cancellationToken = default) {
        ThrowIfDisposed();
        EmailStoreEnumerationOptions effective = options ?? new EmailStoreEnumerationOptions();
        return new EmailStoreAsyncEnumerable<EmailStoreItemReference>(
            token => EnumerateItems(effective, token).GetEnumerator(),
            cancellationToken);
    }

}
