namespace OfficeIMO.Email.Store;

internal sealed class PstStoreSessionBackend : IEmailStoreSessionBackend {
    private readonly PstStoreReader _reader;

    internal PstStoreSessionBackend(Stream stream, EmailStoreFormat format,
        EmailStoreReaderOptions options, CancellationToken cancellationToken) {
        _reader = new PstStoreReader(options);
        _reader.Open(stream, format, loadCompleteIndexes: false, cancellationToken);
    }

    public EmailStoreFormat Format => _reader.Format;
    public string? DisplayName => _reader.DisplayName;
    public long SourceLength => _reader.SourceLength;
    public IReadOnlyList<EmailStoreFolderInfo> Folders => _reader.Folders;
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics => _reader.Diagnostics;

    public IEnumerable<EmailStoreItemReference> EnumerateItems(
        EmailStoreEnumerationOptions options, CancellationToken cancellationToken) =>
        _reader.EnumerateItemReferences(options, cancellationToken);

    public EmailStoreItem ReadItem(EmailStoreItemReference reference,
        CancellationToken cancellationToken) =>
        _reader.ReadReferencedItem(reference, cancellationToken);

    public void Dispose() { }
}
