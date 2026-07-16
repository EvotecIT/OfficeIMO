namespace OfficeIMO.Email.Store;

internal sealed class EmailStoreSessionLifetime : IDisposable {
    private bool _disposed;

    internal void ThrowIfDisposed() {
        if (_disposed) {
            throw new ObjectDisposedException(nameof(EmailStoreSession),
                "Deferred email-store content must be read before its owning session is disposed.");
        }
    }

    public void Dispose() => _disposed = true;
}
