namespace OfficeIMO.Email;

internal sealed partial class EmailSemanticSnapshotBuilder {
    private static readonly HashSet<ushort> PortableExcludedPropertyIds = new HashSet<ushort> {
        0x0039, // submit-time projection is normalized through creation/modified timestamps
        0x0E08, // computed message size
        0x0E09, // parent entry id
        0x0E17, // PST writer materializes a default message-status bitmask
        0x0E21, // attachment ordinal
        0x0FF4, 0x0FF5, 0x0FF6, 0x0FF7, 0x0FF8, 0x0FF9, 0x0FFA, 0x0FFB,
        0x0FFE, 0x0FFF,
        0x3000, 0x300B,
        0x3701, // attachment payload is hashed through the reopenable content abstraction
        0x65E0, 0x65E1, 0x65E2, 0x65E3,
        0x67A4
    };

    private static readonly HashSet<ushort> DeduplicationExcludedPropertyIds =
        new HashSet<ushort>(PortableExcludedPropertyIds) {
            0x007D, // transport-header serialization
            0x0E06, // received timestamp
            0x3007, // creation timestamp
            0x3008  // modification timestamp
        };

    private readonly EmailSemanticComparisonOptions _options;
    private readonly byte[]? _key;
    private readonly Dictionary<string, EmailSemanticEntry> _entries =
        new Dictionary<string, EmailSemanticEntry>(StringComparer.Ordinal);
    private long _attachmentBytesHashed;
    private int _recipientCount;
    private int _attachmentCount;

    internal EmailSemanticSnapshotBuilder(EmailSemanticComparisonOptions options) {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        _key = options.CopyDigestKey();
    }

    internal EmailSemanticSnapshot Build(EmailDocument document, CancellationToken cancellationToken) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        AddDocument(document, "document", 0, useAsync: false, cancellationToken)
            .GetAwaiter().GetResult();
        return Complete();
    }

    internal async Task<EmailSemanticSnapshot> BuildAsync(EmailDocument document,
        CancellationToken cancellationToken) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        await AddDocument(document, "document", 0, useAsync: true, cancellationToken)
            .ConfigureAwait(false);
        return Complete();
    }

    private async Task AddDocument(EmailDocument document, string prefix, int depth,
        bool useAsync, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (depth > _options.MaxEmbeddedMessageDepth) {
            throw new EmailLimitExceededException(
                nameof(EmailSemanticComparisonOptions.MaxEmbeddedMessageDepth),
                depth, _options.MaxEmbeddedMessageDepth);
        }

        AddValue(prefix + "/outlook-item-kind", document.OutlookItemKind);
        AddValue(prefix + "/message-class", document.MessageClass);
        AddValue(prefix + "/body/text", document.Body.Text);
        AddValue(prefix + "/body/html", document.Body.Html);
        AddValue(prefix + "/body/rtf", document.Body.Rtf);
        AddValue(prefix + "/body/text-charset", document.Body.TextCharset);
        AddValue(prefix + "/body/html-charset", document.Body.HtmlCharset);
        AddValue(prefix + "/body/html-content-id", document.Body.HtmlContentId);
        AddValue(prefix + "/body/html-content-location", document.Body.HtmlContentLocation);
        AddValue(prefix + "/body/html-related-root", document.Body.IsHtmlRelatedRoot);
        AddValue(prefix + "/protection/kind", document.Protection.Kind);
        AddValue(prefix + "/protection/message-class", document.Protection.IsProtected
            ? document.Protection.MessageClass
            : null);

        var diagnostics = new List<EmailDiagnostic>();
        MsgPropertyBuilder messageProperties = MsgWriter.CreateMessageProperties(
            document, diagnostics, prefix, EmailWriterOptions.Default,
            document.Format == EmailFileFormat.OutlookTemplate);
        AddMapiProperties(prefix + "/mapi", messageProperties.Properties);

        AddRecipients(document, prefix);
        await AddAttachments(document, prefix, depth, useAsync, cancellationToken)
            .ConfigureAwait(false);

        if (_options.Profile == EmailSemanticComparisonProfile.Strict) {
            AddStrictDocumentEntries(document, prefix);
        }
    }

    private void AddRecipients(EmailDocument document, string prefix) {
        _recipientCount = checked(_recipientCount + document.Recipients.Count);
        for (int index = 0; index < document.Recipients.Count; index++) {
            EmailRecipient recipient = document.Recipients[index];
            string path = string.Concat(prefix, "/recipients/", index.ToString("D8", CultureInfo.InvariantCulture));
            AddValue(path + "/kind", recipient.Kind);
            MsgPropertyBuilder properties = MsgWriter.CreateRecipientProperties(recipient, index);
            AddMapiProperties(path + "/mapi", properties.Properties);
        }
    }

    private void AddStrictDocumentEntries(EmailDocument document, string prefix) {
        AddValue(prefix + "/strict/format", document.Format);
        AddValue(prefix + "/strict/outlook-code-page", document.OutlookCodePage);
        for (int index = 0; index < document.Headers.Count; index++) {
            EmailHeader header = document.Headers[index];
            string path = string.Concat(prefix, "/strict/headers/",
                index.ToString("D8", CultureInfo.InvariantCulture));
            AddValue(path + "/name", header.Name);
            AddValue(path + "/value", header.Value);
            AddValue(path + "/raw-value", header.RawValue);
        }
        AddValue(prefix + "/strict/properties", document.Properties);
        AddTnefAttributes(prefix + "/strict/tnef", document.TnefAttributes);
        AddValue(prefix + "/strict/raw-source", document.RawSource);
    }

    private void AddMapiProperties(string prefix, IEnumerable<MapiProperty> source) {
        var candidates = new List<MapiCandidate>();
        int sourceIndex = 0;
        foreach (MapiProperty property in source) {
            if (!ShouldInclude(property)) { sourceIndex++; continue; }
            string descriptor = Describe(property);
            byte[] valueDigest = EmailSemanticValueDigest.Compute(property.Value, _key);
            candidates.Add(new MapiCandidate(property, descriptor, valueDigest, sourceIndex++));
        }

        candidates.Sort((left, right) => {
            int result = StringComparer.Ordinal.Compare(left.Descriptor, right.Descriptor);
            if (result != 0) return result;
            result = CompareBytes(left.ValueDigest, right.ValueDigest);
            return result != 0 ? result : left.SourceIndex.CompareTo(right.SourceIndex);
        });

        var occurrences = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (MapiCandidate candidate in candidates) {
            occurrences.TryGetValue(candidate.Descriptor, out int occurrence);
            occurrences[candidate.Descriptor] = checked(occurrence + 1);
            string path = string.Concat(prefix, "/", candidate.Descriptor, "/",
                occurrence.ToString("D4", CultureInfo.InvariantCulture));
            AddEntry(path + "/value", candidate.ValueDigest, ValueLength(candidate.Property.Value));
            if (_options.Profile == EmailSemanticComparisonProfile.Strict) {
                AddValue(path + "/flags", candidate.Property.Flags);
                AddValue(path + "/raw-data", candidate.Property.RawData);
            }
        }
    }

    private bool ShouldInclude(MapiProperty property) {
        if (_options.Profile == EmailSemanticComparisonProfile.Strict || property.Name != null) return true;
        HashSet<ushort> excluded = _options.Profile == EmailSemanticComparisonProfile.Deduplication
            ? DeduplicationExcludedPropertyIds
            : PortableExcludedPropertyIds;
        return !excluded.Contains(property.PropertyId);
    }

    private string Describe(MapiProperty property) {
        string type = ((ushort)property.PropertyType).ToString("X4", CultureInfo.InvariantCulture);
        if (property.Name == null) {
            return string.Concat("p-", property.PropertyId.ToString("X4", CultureInfo.InvariantCulture), "-", type);
        }
        string identity = property.Name.Name == null
            ? string.Concat("id-", property.Name.LocalId.GetValueOrDefault().ToString("X8", CultureInfo.InvariantCulture))
            : string.Concat("name-digest-", BitConverter.ToString(
                EmailSemanticValueDigest.Compute(property.Name.Name, _key)).Replace("-", string.Empty));
        return string.Concat("n-", property.Name.PropertySet.ToString("N"), "-", identity, "-", type);
    }

    private void AddTnefAttributes(string prefix, IEnumerable<TnefAttribute> attributes) {
        int index = 0;
        foreach (TnefAttribute attribute in attributes) {
            string path = string.Concat(prefix, "/", index.ToString("D8", CultureInfo.InvariantCulture));
            AddValue(path + "/level", attribute.Level);
            AddValue(path + "/tag", attribute.Tag);
            AddValue(path + "/checksum-valid", attribute.ChecksumIsValid);
            AddValue(path + "/data", attribute.Data);
            index++;
        }
    }

    private void AddValue(string path, object? value) =>
        AddEntry(path, EmailSemanticValueDigest.Compute(value, _key), ValueLength(value));

    private void AddEntry(string path, byte[] digest, long? length) {
        if (_entries.ContainsKey(path)) {
            throw new InvalidDataException(string.Concat("Duplicate semantic path: ", path));
        }
        _entries.Add(path, new EmailSemanticEntry(path, digest, length));
    }

    private EmailSemanticSnapshot Complete() {
        using (var writer = new EmailSemanticHashWriter(_key)) {
            writer.WriteString("OfficeIMO.Email.Semantic");
            writer.WriteInt32(EmailSemanticComparer.CurrentSchemaVersion);
            writer.WriteInt32((int)_options.Profile);
            foreach (EmailSemanticEntry entry in _entries.Values.OrderBy(item => item.Path, StringComparer.Ordinal)) {
                writer.WriteString(entry.Path);
                writer.WriteBytes(entry.Digest);
                writer.WriteInt64(entry.Length ?? -1);
            }
            byte[] digest = writer.Complete();
            var fingerprint = new EmailSemanticFingerprint(EmailSemanticComparer.CurrentSchemaVersion,
                _key == null ? "SHA-256" : "HMAC-SHA-256", digest, _options.Profile,
                _recipientCount, _attachmentCount, _attachmentBytesHashed, _entries.Count);
            return new EmailSemanticSnapshot(fingerprint,
                new Dictionary<string, EmailSemanticEntry>(_entries, StringComparer.Ordinal));
        }
    }

    private static long? ValueLength(object? value) {
        if (value is byte[] bytes) return bytes.LongLength;
        if (value is string text) return text.Length;
        if (value is Array array) return array.LongLength;
        return null;
    }

    private static int CompareBytes(byte[] left, byte[] right) {
        int length = Math.Min(left.Length, right.Length);
        for (int index = 0; index < length; index++) {
            int result = left[index].CompareTo(right[index]);
            if (result != 0) return result;
        }
        return left.Length.CompareTo(right.Length);
    }

    private sealed class MapiCandidate {
        internal MapiCandidate(MapiProperty property, string descriptor, byte[] valueDigest, int sourceIndex) {
            Property = property;
            Descriptor = descriptor;
            ValueDigest = valueDigest;
            SourceIndex = sourceIndex;
        }
        internal MapiProperty Property { get; }
        internal string Descriptor { get; }
        internal byte[] ValueDigest { get; }
        internal int SourceIndex { get; }
    }
}
