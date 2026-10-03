using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Indexes an mbox aggregate with one-message memory and decodes selected entries on demand.</summary>
internal sealed class MboxStoreSessionBackend : IEmailStoreSessionBackend {
    private const string FolderId = "mbox:folder:root";
    private readonly Stream? _stream;
    private readonly Func<Stream>? _openSource;
    private readonly long _sourceLength;
    private readonly EmailStoreReaderOptions _options;
    private readonly EmailStoreDiagnosticCollection _diagnostics = new EmailStoreDiagnosticCollection();
    private readonly EmailStoreReadResources _resources;
    private readonly bool _ownsResources;
    private readonly List<MboxItem> _items = new List<MboxItem>();
    private readonly Dictionary<string, MboxItem> _itemsById =
        new Dictionary<string, MboxItem>(StringComparer.Ordinal);
    private readonly IReadOnlyList<EmailStoreFolderInfo> _folders;

    internal MboxStoreSessionBackend(Stream stream, string? sourceName,
        EmailStoreReaderOptions options, CancellationToken cancellationToken) {
        _stream = stream ?? throw new ArgumentNullException(nameof(stream));
        _options = options ?? throw new ArgumentNullException(nameof(options));
        _resources = new EmailStoreReadResources(options);
        _ownsResources = true;
        _sourceLength = stream.Length;
        DisplayName = GetDisplayName(sourceName);
        Index(stream, cancellationToken);
        _folders = new[] {
            new EmailStoreFolderInfo(FolderId, null, DisplayName ?? "Mailbox", _items.Count, 0)
        };
    }

    internal MboxStoreSessionBackend(Func<Stream> openSource, string? sourceName,
        EmailStoreReaderOptions options, CancellationToken cancellationToken, EmailStoreReadResources? resources = null) {
        _openSource = openSource ?? throw new ArgumentNullException(nameof(openSource));
        _options = options ?? throw new ArgumentNullException(nameof(options));
        _resources = resources ?? new EmailStoreReadResources(options);
        _ownsResources = resources == null;
        DisplayName = GetDisplayName(sourceName);
        using (Stream stream = openSource()) {
            _sourceLength = stream.Length;
            Index(stream, cancellationToken);
        }
        _folders = new[] {
            new EmailStoreFolderInfo(FolderId, null, DisplayName ?? "Mailbox", _items.Count, 0)
        };
    }

    public EmailStoreFormat Format => EmailStoreFormat.Mbox;
    public string? DisplayName { get; }
    public long SourceLength => _sourceLength;
    public IReadOnlyList<EmailStoreFolderInfo> Folders => _folders;
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics => _diagnostics;

    public IEnumerable<EmailStoreItemReference> EnumerateItems(
        EmailStoreEnumerationOptions options, CancellationToken cancellationToken) {
        if (options.FolderId != null && !string.Equals(options.FolderId, FolderId,
            StringComparison.Ordinal)) {
            throw new KeyNotFoundException("The requested folder does not belong to this mbox session.");
        }
        if (!options.IncludeRegularItems) yield break;
        int count = 0;
        foreach (MboxItem item in _items) {
            cancellationToken.ThrowIfCancellationRequested();
            if (++count > options.MaxItems) yield break;
            yield return new EmailStoreItemReference(
                item.Id, FolderId, false, false, item.Summary);
        }
    }

    public EmailStoreItemSummary ReadSummary(EmailStoreItemReference reference,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        return GetItem(reference).Summary;
    }

    public EmailStoreItem ReadItem(EmailStoreItemReference reference, EmailStoreItemReadOptions options,
        CancellationToken cancellationToken) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        cancellationToken.ThrowIfCancellationRequested();
        MboxItem item = GetItem(reference);
        EmailMailboxEntryReadResult entry;
        try {
            var mailboxOptions = CreateMailboxOptions(
                EmailStoreMessageReader.CreateOptions(_options,
                    includeAttachmentContent: options.Includes(EmailStoreItemReadParts.AttachmentContent),
                    maxDecodedPropertyBytes: options.MaxDecodedPropertyBytes,
                    includeEmbeddedMessages: options.Includes(EmailStoreItemReadParts.EmbeddedItems)),
                maximumMessages: 1);
            Stream source = _openSource?.Invoke() ?? _stream!;
            try {
                if (source.Length != _sourceLength) {
                    throw new InvalidDataException("The mbox source length changed after it was indexed.");
                }
                using (var input = new EmailStoreSegmentStream(source, item.Offset, item.Length)) {
                    if (options.PreferStreamingAttachmentContent) {
                        using EmailReadResult result = MboxSelectedMessageReader.Read(input, mailboxOptions,
                            cancellationToken, out EmailMailboxEntry mailboxEntry);
                        _resources.Adopt(result);
                        entry = new EmailMailboxEntryReadResult(mailboxEntry, result.Diagnostics, item.Length);
                    } else entry = new EmailMailboxReader(mailboxOptions).ReadEntries(input, cancellationToken).Single();
                }
            } finally {
                if (_openSource != null) source.Dispose();
            }
        } catch (EmailLimitExceededException exception) {
            throw ConvertLimit(exception);
        }
        AddDiagnostics(entry.Diagnostics, item.Id);
        EmailDocument document = entry.Entry.Document;
        ApplyStoreProperties(document, item.Id, entry.Entry);
        EmailStoreItemReadParts loadedParts = EmailStoreItemReadParts.All;
        if (!options.Includes(EmailStoreItemReadParts.AttachmentContent)) {
            loadedParts &= ~EmailStoreItemReadParts.AttachmentContent;
        }
        if (!options.Includes(EmailStoreItemReadParts.EmbeddedItems)) loadedParts &= ~EmailStoreItemReadParts.EmbeddedItems;
        return new EmailStoreItem(item.Id, FolderId, document,
            loadedParts: loadedParts, format: EmailStoreFormat.Mbox, summary: item.Summary);
    }

    public void Dispose() { if (_ownsResources) _resources.Dispose(); }

    private void Index(Stream stream, CancellationToken cancellationToken) {
        long offset = 0;
        long totalAttachmentBytes = 0;
        var mailboxOptions = CreateMailboxOptions(
            EmailStoreMessageReader.CreateOptions(_options, includeAttachmentContent: false, includeEmbeddedMessages: false),
            _options.MaxItemCount);
        try {
            int index = 0;
            foreach (EmailMailboxEntryReadResult result in
                     new EmailMailboxReader(mailboxOptions).ReadEntries(stream, cancellationToken)) {
                cancellationToken.ThrowIfCancellationRequested();
                totalAttachmentBytes = EmailStoreAttachmentBudget.AddDocument(
                    result.Entry.Document, totalAttachmentBytes, _options.MaxTotalAttachmentBytes);
                string id = "mbox:item:" + index.ToString("D8", CultureInfo.InvariantCulture);
                var summary = new EmailStoreItemSummary(
                    result.Entry.Document,
                    result.Entry.Document.Attachments.Count > 0,
                    result.Entry.Document.MessageMetadata.IsRead);
                var item = new MboxItem(id, offset, result.BytesRead, summary);
                _items.Add(item);
                _itemsById.Add(id, item);
                AddDiagnostics(result.Diagnostics, id);
                offset = checked(offset + result.BytesRead);
                index++;
            }
        } catch (EmailLimitExceededException exception) {
            throw ConvertLimit(exception);
        } catch (InvalidDataException exception) when (
            exception.Message.StartsWith("EMAIL_MBOX_ENVELOPE_MISSING", StringComparison.Ordinal)) {
            _diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_MBOX_ENVELOPE_MISSING",
                "The mailbox does not begin with an mbox From separator.",
                EmailStoreDiagnosticSeverity.Error));
            offset = stream.Length;
        }
        if (offset != stream.Length) {
            throw new InvalidDataException("The mbox index did not consume the complete source stream.");
        }
    }

    private EmailMailboxReaderOptions CreateMailboxOptions(
        EmailReaderOptions messageOptions, int maximumMessages) =>
        new EmailMailboxReaderOptions(
            _options.MaxInputBytes, messageOptions, MboxVariant.Auto, maximumMessages);

    private MboxItem GetItem(EmailStoreItemReference reference) {
        if (!_itemsById.TryGetValue(reference.Id, out MboxItem? item) ||
            !string.Equals(reference.FolderId, FolderId, StringComparison.Ordinal) ||
            reference.IsAssociated || reference.IsOrphaned) {
            throw new KeyNotFoundException("The item reference does not belong to this mbox session.");
        }
        return item;
    }

    private void AddDiagnostics(IEnumerable<EmailDiagnostic> diagnostics, string itemId) {
        foreach (EmailDiagnostic diagnostic in diagnostics) {
            EmailStoreDiagnosticSeverity severity = diagnostic.Severity == EmailDiagnosticSeverity.Error
                ? EmailStoreDiagnosticSeverity.Error
                : diagnostic.Severity == EmailDiagnosticSeverity.Information
                    ? EmailStoreDiagnosticSeverity.Information
                    : EmailStoreDiagnosticSeverity.Warning;
            string location = diagnostic.Location == null
                ? itemId
                : string.Concat(itemId, "/", diagnostic.Location);
            _diagnostics.Add(new EmailStoreDiagnostic(
                diagnostic.Code, diagnostic.Message, severity, location,
                diagnostic.Operation, diagnostic.ByteOffset, diagnostic.LimitName,
                diagnostic.ActualValue, diagnostic.MaximumValue, diagnostic.Disposition,
                diagnostic.DataLossRisk, diagnostic.SuggestedAction, diagnostic.IsRetryable));
        }
    }

    private static void ApplyStoreProperties(EmailDocument document, string itemId,
        EmailMailboxEntry entry) {
        document.Properties["EmailStore:Format"] = EmailStoreFormat.Mbox.ToString();
        document.Properties["EmailStore:ItemId"] = itemId;
        document.Properties["EmailStore:FolderId"] = FolderId;
        if (entry.EnvelopeSender != null) document.Properties["Mbox:EnvelopeSender"] = entry.EnvelopeSender;
        if (entry.EnvelopeDate.HasValue) document.Properties["Mbox:EnvelopeDate"] = entry.EnvelopeDate.Value;
        if (entry.RawFromLine != null) document.Properties["Mbox:RawFromLine"] = entry.RawFromLine;
    }

    private static string? GetDisplayName(string? sourceName) {
        if (string.IsNullOrWhiteSpace(sourceName)) return "Mailbox";
        try { return Path.GetFileNameWithoutExtension(sourceName); }
        catch (Exception exception) when (exception is ArgumentException || exception is NotSupportedException) {
            return sourceName;
        }
    }

    private static EmailStoreLimitExceededException ConvertLimit(EmailLimitExceededException exception) {
        string name = exception.LimitName == nameof(EmailMailboxReaderOptions.MaxMailboxBytes)
            ? nameof(EmailStoreReaderOptions.MaxInputBytes)
            : exception.LimitName == nameof(EmailMailboxReaderOptions.MaxMessageCount)
                ? nameof(EmailStoreReaderOptions.MaxItemCount)
                : exception.LimitName == nameof(EmailReaderOptions.MaxInputBytes)
                    ? nameof(EmailStoreReaderOptions.MaxMessageBytes)
                    : exception.LimitName == nameof(EmailReaderOptions.MaxAttachmentBytes)
                        ? nameof(EmailStoreReaderOptions.MaxAttachmentBytes)
                        : exception.LimitName == nameof(EmailReaderOptions.MaxTotalAttachmentBytes)
                            ? nameof(EmailStoreReaderOptions.MaxTotalAttachmentBytes)
                            : exception.LimitName == nameof(EmailReaderOptions.MaxDecodedPropertyBytes)
                                ? nameof(EmailStoreReaderOptions.MaxDecodedPropertyBytesPerItem)
                            : exception.LimitName;
        return new EmailStoreLimitExceededException(name, exception.ActualValue, exception.MaximumValue);
    }

    private sealed class MboxItem {
        internal MboxItem(string id, long offset, long length, EmailStoreItemSummary summary) {
            Id = id;
            Offset = offset;
            Length = length;
            Summary = summary;
        }
        internal string Id { get; }
        internal long Offset { get; }
        internal long Length { get; }
        internal EmailStoreItemSummary Summary { get; }
    }

}
