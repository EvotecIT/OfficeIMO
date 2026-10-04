namespace OfficeIMO.Email.Store;

/// <summary>Bounds and owns temporary attachment content for the lifetime of one store session.</summary>
internal sealed class EmailStoreReadResources : IDisposable {
    private readonly EmailStoreReaderOptions _options;
    private readonly List<IDisposable> _resources = new List<IDisposable>();
    private long _attachmentBytes;
    private bool _disposed;

    internal EmailStoreReadResources(EmailStoreReaderOptions options) => _options = options;

    internal void Adopt(EmailReadResult result) {
        if (_disposed) throw new ObjectDisposedException(nameof(EmailStoreSession));
        if (!result.UsesFileBackedContent) return;
        if (_resources.Count >= _options.MaxItemCount) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxItemCount),
                _resources.Count + 1L, _options.MaxItemCount);
        }
        long total = EmailStoreAttachmentBudget.AddDocument(result.Document, _attachmentBytes, _options.MaxTotalAttachmentBytes);
        IDisposable? resource = result.DetachResources();
        if (resource != null) _resources.Add(resource);
        _attachmentBytes = total;
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        foreach (IDisposable resource in _resources) resource.Dispose();
        _resources.Clear();
    }
}
