namespace OfficeIMO.Email.Store;

internal sealed class PstStoreSessionBackend : IEmailStoreSessionBackend {
    private readonly EmailStoreSessionLifetime _lifetime = new EmailStoreSessionLifetime();
    private readonly PstStoreReader _reader;

    internal PstStoreSessionBackend(Stream stream, EmailStoreFormat format,
        EmailStoreReaderOptions options, CancellationToken cancellationToken) {
        _reader = new PstStoreReader(options, _lifetime);
        _reader.Open(stream, format, loadCompleteIndexes: false, cancellationToken);
    }

    public EmailStoreFormat Format => _reader.Format;
    public string? DisplayName => _reader.DisplayName;
    public long SourceLength => _reader.SourceLength;
    public IReadOnlyList<EmailStoreFolderInfo> Folders => _reader.Folders;
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics => _reader.Diagnostics;
    internal bool IsPasswordProtected => _reader.IsPasswordProtected;

    public IEnumerable<EmailStoreItemReference> EnumerateItems(
        EmailStoreEnumerationOptions options, CancellationToken cancellationToken) =>
        _reader.EnumerateItemReferences(options, cancellationToken);

    public EmailStoreItemSummary ReadSummary(EmailStoreItemReference reference,
        CancellationToken cancellationToken) =>
        _reader.ReadReferencedSummary(reference, cancellationToken);

    public EmailStoreItem ReadItem(EmailStoreItemReference reference, EmailStoreItemReadOptions options,
        CancellationToken cancellationToken) =>
        _reader.ReadReferencedItem(reference, options, cancellationToken);

    internal EmailStoreStructuralValidationResult ValidateStructure(
        EmailStoreValidationOptions options, CancellationToken cancellationToken) =>
        _reader.ValidateStructure(options, cancellationToken);

    public void Dispose() => _lifetime.Dispose();
}
