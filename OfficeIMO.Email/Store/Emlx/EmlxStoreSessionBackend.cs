namespace OfficeIMO.Email.Store;

/// <summary>Projects one EMLX item on demand and owns explicitly requested streaming attachment content.</summary>
internal sealed class EmlxStoreSessionBackend : IEmailStoreSessionBackend {
    private readonly Stream _stream;
    private readonly string? _sourceName;
    private readonly EmailStoreReaderOptions _options;
    private readonly EmailStoreReadResources _resources;
    private readonly EmailStoreItemReference _reference;
    private readonly EmailStoreItemSummary _summary;
    private readonly string? _sourceFingerprint;
    private bool _sourceChanged;
    private readonly EmailStoreDiagnosticCollection _diagnostics = new EmailStoreDiagnosticCollection();

    internal EmlxStoreSessionBackend(Stream stream, string? sourceName, EmailStoreReaderOptions options,
        CancellationToken cancellationToken, bool isSnapshot = false) {
        _stream = stream;
        _sourceName = sourceName;
        _options = options;
        _resources = new EmailStoreReadResources(options);
        SourceLength = stream.Length;
        _sourceFingerprint = isSnapshot ? null : EmailStoreSourceFingerprint.Compute(stream, SourceLength, options.MaxInputBytes, cancellationToken);
        EmailStoreReadResult index = new EmlxStoreReader(options, includeAttachmentContent: false,
            includeEmbeddedMessages: false).Read(stream, sourceName, cancellationToken);
        DisplayName = index.Store.DisplayName;
        EmailStoreFolder folder = index.Store.Folders.Single();
        EmailStoreItem item = folder.Items.Single();
        _summary = EmailStoreItemSummary.FromItem(item);
        _reference = new EmailStoreItemReference(item.Id, folder.Id, false, false, _summary);
        Folders = new[] { new EmailStoreFolderInfo(folder.Id, null, folder.Name, 1, 0) };
        AddDiagnostics(index.Diagnostics);
        ValidateSource(cancellationToken);
    }

    public EmailStoreFormat Format => EmailStoreFormat.Emlx;
    public string? DisplayName { get; }
    public long SourceLength { get; }
    public IReadOnlyList<EmailStoreFolderInfo> Folders { get; }
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics => _diagnostics;

    public IEnumerable<EmailStoreItemReference> EnumerateItems(EmailStoreEnumerationOptions options,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (options.FolderId != null && options.FolderId != _reference.FolderId) {
            throw new KeyNotFoundException("The requested folder does not belong to this EMLX session.");
        }
        if (options.IncludeRegularItems) yield return _reference;
    }

    public EmailStoreItemSummary ReadSummary(EmailStoreItemReference reference, CancellationToken cancellationToken) {
        ValidateReference(reference, cancellationToken);
        return _summary;
    }

    public EmailStoreItem ReadItem(EmailStoreItemReference reference, EmailStoreItemReadOptions options,
        CancellationToken cancellationToken) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        ValidateReference(reference, cancellationToken);
        bool includeContent = options.Includes(EmailStoreItemReadParts.AttachmentContent);
        EmailStoreReadResult result = new EmlxStoreReader(_options, includeContent, options.MaxDecodedPropertyBytes,
            options.Includes(EmailStoreItemReadParts.EmbeddedItems), _resources, options.PreferStreamingAttachmentContent)
            .Read(_stream, _sourceName, cancellationToken);
        ValidateSource(cancellationToken);
        AddDiagnostics(result.Diagnostics);
        return result.Store.Folders.Single().Items.Single();
    }

    private void ValidateReference(EmailStoreItemReference reference, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (reference.Id != _reference.Id || reference.FolderId != _reference.FolderId || reference.IsAssociated || reference.IsOrphaned) {
            throw new KeyNotFoundException("The item reference does not belong to this EMLX session.");
        }
        ValidateSource(cancellationToken);
    }

    private void ValidateSource(CancellationToken cancellationToken) {
        if (_sourceChanged) throw new InvalidDataException("The EMLX source changed after it was indexed.");
        try {
            if (_stream.Length != SourceLength || _sourceFingerprint != null &&
                !string.Equals(_sourceFingerprint, EmailStoreSourceFingerprint.Compute(_stream, SourceLength,
                    _options.MaxInputBytes, cancellationToken), StringComparison.Ordinal)) {
                throw new InvalidDataException("The EMLX source changed after it was indexed.");
            }
        } catch (InvalidDataException) {
            _sourceChanged = true;
            _resources.Dispose();
            throw;
        }
    }

    private void AddDiagnostics(IEnumerable<EmailStoreDiagnostic> diagnostics) {
        foreach (EmailStoreDiagnostic diagnostic in diagnostics) _diagnostics.Add(diagnostic);
    }

    public void Dispose() => _resources.Dispose();
}
