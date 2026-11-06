namespace OfficeIMO.Email.Store;

/// <summary>Exposes active messages in one DBX message database through the common store session.</summary>
internal sealed class DbxStoreSessionBackend : IEmailStoreSessionBackend {
    private const string FolderId = "dbx:folder:root";
    private const string MessageIdPrefix = "dbx:item:";
    private const uint ReadFlag = 0x80;
    private const uint AttachmentFlag = 0x4000;

    private readonly Stream? _source;
    private readonly Func<Stream>? _openSource;
    private readonly EmailStoreReaderOptions _options;
    private readonly EmailStoreSourceGuard _guard;
    private readonly EmailStoreReadResources _resources;
    private readonly bool _ownsResources;
    private readonly List<uint> _records;
    private readonly HashSet<uint> _recordAddresses;
    private readonly EmailStoreDiagnosticCollection _diagnostics = new EmailStoreDiagnosticCollection();

    internal DbxStoreSessionBackend(Stream source, string? sourceName, EmailStoreReaderOptions options,
        CancellationToken cancellationToken, bool isSnapshot = false)
        : this(source, null, sourceName, options, cancellationToken, null, 0, isSnapshot) { }

    internal DbxStoreSessionBackend(Stream input, Func<Stream> openSource, string? sourceName,
        EmailStoreReaderOptions options, CancellationToken cancellationToken,
        EmailStoreReadResources resources, int indexedItemCount)
        : this(input, openSource, sourceName, options, cancellationToken, resources, indexedItemCount,
            isSnapshot: false) { }

    private DbxStoreSessionBackend(Stream input, Func<Stream>? openSource, string? sourceName,
        EmailStoreReaderOptions options, CancellationToken cancellationToken,
        EmailStoreReadResources? resources, int indexedItemCount, bool isSnapshot) {
        _source = openSource == null ? input : null;
        _openSource = openSource;
        _options = options;
        _resources = resources ?? new EmailStoreReadResources(options);
        _ownsResources = resources == null;
        DisplayName = string.IsNullOrWhiteSpace(sourceName)
            ? "Mailbox" : Path.GetFileNameWithoutExtension(sourceName) ?? "Mailbox";
        try {
            SourceLength = input.Length;
            _guard = new EmailStoreSourceGuard(input, SourceLength, options.MaxInputBytes, _resources.Dispose,
                cancellationToken, isSnapshot);
            _records = new DbxDatabase(input, options).ReadIndex(folderDatabase: false, cancellationToken,
                indexedItemCount);
            _recordAddresses = new HashSet<uint>(_records);
            _guard.Validate(input, cancellationToken);
        } catch {
            if (_ownsResources) _resources.Dispose();
            throw;
        }
        Folders = new[] { new EmailStoreFolderInfo(FolderId, null, DisplayName, _records.Count, 0) };
    }

    public EmailStoreFormat Format => EmailStoreFormat.Dbx;
    public string DisplayName { get; }
    public long SourceLength { get; }
    public IReadOnlyList<EmailStoreFolderInfo> Folders { get; }
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics => _diagnostics;

    public IEnumerable<EmailStoreItemReference> EnumerateItems(EmailStoreEnumerationOptions options,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (options.FolderId != null && !string.Equals(options.FolderId, FolderId, StringComparison.Ordinal))
            throw new KeyNotFoundException("The requested folder does not belong to this DBX session.");
        if (!options.IncludeRegularItems) yield break;
        int count = Math.Min(options.MaxItems, _records.Count);
        for (int index = 0; index < count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            string id = MessageIdPrefix + _records[index].ToString("x8", CultureInfo.InvariantCulture);
            yield return new EmailStoreItemReference(id, FolderId, isAssociated: false, isOrphaned: false);
        }
    }

    public EmailStoreItemSummary ReadSummary(EmailStoreItemReference reference, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        uint address = GetAddress(reference);
        using Stream? owned = _openSource?.Invoke();
        Stream input = owned ?? _source!;
        using Stream message = OpenIndexedMessage(input, address, cancellationToken, out uint? flags);
        var diagnostics = new List<EmailDiagnostic>();
        EmailDocument document;
        try {
            document = MimeHeaderParser.ReadEnvelope(message, EmailStoreMessageReader.CreateOptions(_options),
                diagnostics, cancellationToken);
        } catch (EmailLimitExceededException exception) {
            throw EmailStoreMessageReader.ConvertLimit(exception);
        }
        _guard.Validate(input, cancellationToken);
        CopyDiagnostics(diagnostics, reference.Id);
        bool? hasAttachments = flags.HasValue ? (flags.Value & AttachmentFlag) != 0 : null;
        bool? isRead = flags.HasValue ? (flags.Value & ReadFlag) != 0 : document.MessageMetadata.IsRead;
        return new EmailStoreItemSummary(document, hasAttachments, isRead, isHeaderOnly: false);
    }

    public EmailStoreItem ReadItem(EmailStoreItemReference reference, EmailStoreItemReadOptions options,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        uint address = GetAddress(reference);
        using Stream? owned = _openSource?.Invoke();
        return ReadItem(owned ?? _source!, address, reference, options, cancellationToken);
    }

    private EmailStoreItem ReadItem(Stream input, uint address, EmailStoreItemReference reference,
        EmailStoreItemReadOptions options, CancellationToken cancellationToken) {
        using Stream message = OpenIndexedMessage(input, address, cancellationToken, out uint? flags);
        bool includeAttachments = options.Includes(EmailStoreItemReadParts.AttachmentContent);
        bool includeEmbedded = options.Includes(EmailStoreItemReadParts.EmbeddedItems);

        using EmailReadResult result = EmailStoreMessageReader.Read(message, _options, cancellationToken,
            includeAttachments, options.MaxDecodedPropertyBytes, includeEmbedded, options.PreferStreamingAttachmentContent);
        if (result.Document.Format != EmailFileFormat.Eml)
            throw DbxDatabase.Invalid("The indexed payload is not an RFC message.");
        _guard.Validate(input, cancellationToken);
        _resources.Adopt(result);
        CopyDiagnostics(result.Diagnostics, reference.Id);

        EmailDocument document = result.Document;
        document.Properties["EmailStore:Format"] = Format.ToString();
        document.Properties["EmailStore:FolderId"] = FolderId;
        document.Properties["EmailStore:ItemId"] = reference.Id;
        if (flags.HasValue) {
            document.MessageMetadata.IsRead = (flags.Value & ReadFlag) != 0;
            document.MessageMetadata.IsDraft = (flags.Value & 0x08) != 0;
        }

        EmailStoreItemReadParts loaded = EmailStoreItemReadParts.All;
        if (!includeAttachments) loaded &= ~EmailStoreItemReadParts.AttachmentContent;
        if (!includeEmbedded) loaded &= ~EmailStoreItemReadParts.EmbeddedItems;
        return new EmailStoreItem(reference.Id, FolderId, document, loadedParts: loaded, format: Format);
    }

    private Stream OpenIndexedMessage(Stream input, uint address, CancellationToken cancellationToken, out uint? flags) {
        _guard.Validate(input, cancellationToken);
        var database = new DbxDatabase(input, _options);
        DbxIndexedRecord record = database.ReadIndexedRecord(address, cancellationToken);
        flags = record.ReadUInt32(0x01, cancellationToken);
        uint firstFragment = record.ReadUInt32(0x04, cancellationToken) ?? 0;
        return database.OpenMessage(firstFragment, cancellationToken);
    }

    private void CopyDiagnostics(IEnumerable<EmailDiagnostic> diagnostics, string itemId) {
        foreach (EmailDiagnostic diagnostic in diagnostics) {
            EmailStoreDiagnosticSeverity severity = diagnostic.Severity switch {
                EmailDiagnosticSeverity.Error => EmailStoreDiagnosticSeverity.Error,
                EmailDiagnosticSeverity.Warning => EmailStoreDiagnosticSeverity.Warning,
                _ => EmailStoreDiagnosticSeverity.Information
            };
            string location = diagnostic.Location == null ? itemId : itemId + "/" + diagnostic.Location;
            _diagnostics.Add(new EmailStoreDiagnostic(diagnostic.Code, diagnostic.Message, severity, location,
                diagnostic.Operation, diagnostic.ByteOffset, diagnostic.LimitName, diagnostic.ActualValue,
                diagnostic.MaximumValue, diagnostic.Disposition, diagnostic.DataLossRisk, diagnostic.SuggestedAction,
                diagnostic.IsRetryable));
        }
    }

    public void Dispose() { if (_ownsResources) _resources.Dispose(); }

    private uint GetAddress(EmailStoreItemReference reference) {
        if (!string.Equals(reference.FolderId, FolderId, StringComparison.Ordinal) ||
            reference.IsAssociated || reference.IsOrphaned ||
            !reference.Id.StartsWith(MessageIdPrefix, StringComparison.Ordinal) ||
            reference.Id.Length != MessageIdPrefix.Length + 8 ||
            !uint.TryParse(reference.Id.Substring(MessageIdPrefix.Length), NumberStyles.AllowHexSpecifier,
                CultureInfo.InvariantCulture, out uint address) || !_recordAddresses.Contains(address)) {
            throw new KeyNotFoundException("The item reference does not belong to this DBX session.");
        }
        return address;
    }
}
