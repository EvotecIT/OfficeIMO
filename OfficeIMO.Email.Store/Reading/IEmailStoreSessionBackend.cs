namespace OfficeIMO.Email.Store;

internal interface IEmailStoreSessionBackend : IDisposable {
    EmailStoreFormat Format { get; }
    string? DisplayName { get; }
    long SourceLength { get; }
    IReadOnlyList<EmailStoreFolderInfo> Folders { get; }
    IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }
    IEnumerable<EmailStoreItemReference> EnumerateItems(
        EmailStoreEnumerationOptions options, CancellationToken cancellationToken);
    EmailStoreItem ReadItem(EmailStoreItemReference reference, CancellationToken cancellationToken);
}
