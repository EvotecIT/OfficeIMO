namespace OfficeIMO.Email.Store;

internal interface IEmailStoreSessionBackend : IDisposable {
    EmailStoreFormat Format { get; }
    string? DisplayName { get; }
    long SourceLength { get; }
    IReadOnlyList<EmailStoreFolderInfo> Folders { get; }
    IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }
    IEnumerable<EmailStoreItemReference> EnumerateItems(
        EmailStoreEnumerationOptions options, CancellationToken cancellationToken);
    EmailStoreItemSummary ReadSummary(EmailStoreItemReference reference, CancellationToken cancellationToken);
    EmailStoreItem ReadItem(EmailStoreItemReference reference, EmailStoreItemReadOptions options,
        CancellationToken cancellationToken);
}
