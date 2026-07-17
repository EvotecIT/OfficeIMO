using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Indexes an mbox aggregate with one-message memory and decodes selected entries on demand.</summary>
internal sealed class MboxStoreSessionBackend : IEmailStoreSessionBackend {
    private const string FolderId = "mbox:folder:root";
    private readonly Stream _stream;
    private readonly EmailStoreReaderOptions _options;
    private readonly List<EmailStoreDiagnostic> _diagnostics = new List<EmailStoreDiagnostic>();
    private readonly HashSet<DiagnosticKey> _diagnosticKeys = new HashSet<DiagnosticKey>();
    private readonly List<MboxItem> _items = new List<MboxItem>();
    private readonly Dictionary<string, MboxItem> _itemsById =
        new Dictionary<string, MboxItem>(StringComparer.Ordinal);
    private readonly IReadOnlyList<EmailStoreFolderInfo> _folders;

    internal MboxStoreSessionBackend(Stream stream, string? sourceName,
        EmailStoreReaderOptions options, CancellationToken cancellationToken) {
        _stream = stream ?? throw new ArgumentNullException(nameof(stream));
        _options = options ?? throw new ArgumentNullException(nameof(options));
        DisplayName = GetDisplayName(sourceName);
        Index(cancellationToken);
        _folders = new[] {
            new EmailStoreFolderInfo(FolderId, null, DisplayName ?? "Mailbox", _items.Count, 0)
        };
    }

    public EmailStoreFormat Format => EmailStoreFormat.Mbox;
    public string? DisplayName { get; }
    public long SourceLength => _stream.Length;
    public IReadOnlyList<EmailStoreFolderInfo> Folders => _folders;
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics => _diagnostics;

    public IEnumerable<EmailStoreItemReference> EnumerateItems(
        EmailStoreEnumerationOptions options, CancellationToken cancellationToken) {
        if (options.FolderId != null && !string.Equals(options.FolderId, FolderId,
            StringComparison.Ordinal)) {
            throw new KeyNotFoundException("The requested folder does not belong to this mbox session.");
        }
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
                EmailStoreMessageReader.CreateOptions(_options), maximumMessages: 1);
            using (var input = new ReadOnlySegmentStream(_stream, item.Offset, item.Length)) {
                entry = new EmailMailboxReader(mailboxOptions).ReadEntries(input, cancellationToken).Single();
            }
        } catch (EmailLimitExceededException exception) {
            throw ConvertLimit(exception);
        }
        AddDiagnostics(entry.Diagnostics, item.Id);
        EmailDocument document = entry.Entry.Document;
        ApplyStoreProperties(document, item.Id, entry.Entry);
        EmailStoreItemReadParts loadedParts = _options.RetainAttachmentContent
            ? EmailStoreItemReadParts.All
            : EmailStoreItemReadParts.All & ~EmailStoreItemReadParts.AttachmentContent;
        return new EmailStoreItem(item.Id, FolderId, document,
            loadedParts: loadedParts, format: EmailStoreFormat.Mbox, summary: item.Summary);
    }

    public void Dispose() { }

    private void Index(CancellationToken cancellationToken) {
        long offset = 0;
        long totalAttachmentBytes = 0;
        var mailboxOptions = CreateMailboxOptions(
            EmailStoreMessageReader.CreateOptions(_options, includeAttachmentContent: false),
            _options.MaxItemCount);
        try {
            int index = 0;
            foreach (EmailMailboxEntryReadResult result in
                     new EmailMailboxReader(mailboxOptions).ReadEntries(_stream, cancellationToken)) {
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
            offset = _stream.Length;
        }
        if (offset != _stream.Length) {
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
            var key = new DiagnosticKey(diagnostic.Code, diagnostic.Message, severity, location);
            if (_diagnosticKeys.Add(key)) {
                _diagnostics.Add(new EmailStoreDiagnostic(
                    diagnostic.Code, diagnostic.Message, severity, location));
            }
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

    private readonly struct DiagnosticKey : IEquatable<DiagnosticKey> {
        internal DiagnosticKey(string code, string message,
            EmailStoreDiagnosticSeverity severity, string location) {
            Code = code;
            Message = message;
            Severity = severity;
            Location = location;
        }

        private string Code { get; }
        private string Message { get; }
        private EmailStoreDiagnosticSeverity Severity { get; }
        private string Location { get; }

        public bool Equals(DiagnosticKey other) =>
            string.Equals(Code, other.Code, StringComparison.Ordinal) &&
            string.Equals(Message, other.Message, StringComparison.Ordinal) &&
            Severity == other.Severity &&
            string.Equals(Location, other.Location, StringComparison.Ordinal);

        public override bool Equals(object? obj) => obj is DiagnosticKey other && Equals(other);

        public override int GetHashCode() {
            unchecked {
                int hash = StringComparer.Ordinal.GetHashCode(Code);
                hash = hash * 397 ^ StringComparer.Ordinal.GetHashCode(Message);
                hash = hash * 397 ^ (int)Severity;
                return hash * 397 ^ StringComparer.Ordinal.GetHashCode(Location);
            }
        }
    }

    private sealed class ReadOnlySegmentStream : Stream {
        private readonly Stream _source;
        private readonly long _start;
        private readonly long _length;
        private long _position;

        internal ReadOnlySegmentStream(Stream source, long start, long length) {
            _source = source;
            _start = start;
            _length = length;
        }

        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => _length;
        public override long Position {
            get => _position;
            set {
                if (value < 0 || value > _length) throw new ArgumentOutOfRangeException(nameof(value));
                _position = value;
            }
        }

        public override int Read(byte[] buffer, int offset, int count) {
            if (buffer == null) throw new ArgumentNullException(nameof(buffer));
            if (offset < 0 || count < 0 || offset > buffer.Length - count)
                throw new ArgumentOutOfRangeException();
            int bounded = (int)Math.Min(count, _length - _position);
            if (bounded == 0) return 0;
            long absolutePosition = checked(_start + _position);
            if (_source.Position != absolutePosition) _source.Position = absolutePosition;
            int read = _source.Read(buffer, offset, bounded);
            _position += read;
            return read;
        }

        public override int ReadByte() {
            if (_position >= _length) return -1;
            long absolutePosition = checked(_start + _position);
            if (_source.Position != absolutePosition) _source.Position = absolutePosition;
            int value = _source.ReadByte();
            if (value >= 0) _position++;
            return value;
        }

        public override long Seek(long offset, SeekOrigin origin) {
            long target = origin == SeekOrigin.Begin
                ? offset
                : origin == SeekOrigin.Current
                    ? checked(_position + offset)
                    : checked(_length + offset);
            Position = target;
            return _position;
        }

        public override void Flush() { }
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
