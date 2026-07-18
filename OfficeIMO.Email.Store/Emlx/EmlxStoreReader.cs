using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Reads one Apple Mail EMLX envelope and delegates its RFC message to OfficeIMO.Email.</summary>
internal sealed class EmlxStoreReader {
    private const int MaxLengthPrefixBytes = 64;
    private const string FolderId = "emlx:folder:apple-mail";
    private readonly EmailStoreReaderOptions _options;
    private readonly bool? _includeAttachmentContent;
    private readonly List<EmailStoreDiagnostic> _diagnostics = new List<EmailStoreDiagnostic>();

    internal EmlxStoreReader(EmailStoreReaderOptions options, bool? includeAttachmentContent = null) {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        _includeAttachmentContent = includeAttachmentContent;
    }

    internal static bool HasEnvelopePrefix(Stream stream) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead || !stream.CanSeek) return false;
        long originalPosition = stream.Position;
        try {
            stream.Position = 0;
            long declaredMessageBytes = ReadLengthPrefix(stream, CancellationToken.None);
            return declaredMessageBytes <= stream.Length - stream.Position;
        } catch (Exception exception) when (exception is InvalidDataException || exception is IOException ||
                                             exception is NotSupportedException) {
            return false;
        } finally {
            stream.Position = originalPosition;
        }
    }

    internal EmailStoreReadResult Read(Stream stream, string? sourceName, CancellationToken cancellationToken) {
        stream.Position = 0;
        long declaredMessageBytes = ReadLengthPrefix(stream, cancellationToken);
        if (declaredMessageBytes > _options.MaxMessageBytes) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxMessageBytes),
                declaredMessageBytes, _options.MaxMessageBytes);
        }
        if (declaredMessageBytes > int.MaxValue) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxMessageBytes),
                declaredMessageBytes, Math.Min(_options.MaxMessageBytes, int.MaxValue));
        }
        if (declaredMessageBytes > stream.Length - stream.Position) {
            throw new InvalidDataException("The EMLX message byte count exceeds the remaining source length.");
        }

        byte[] messageBytes = ReadExact(stream, (int)declaredMessageBytes, cancellationToken,
            "The EMLX message ended before its declared byte count.");
        EmailReadResult emailResult = ReadMessage(messageBytes, cancellationToken);
        CopyDiagnostics(emailResult.Diagnostics);

        string itemName = GetItemName(sourceName);
        string itemId = string.Concat("emlx:item:", itemName);
        var store = new EmailStore {
            Format = EmailStoreFormat.Emlx,
            DisplayName = GetDisplayName(itemName)
        };
        var folder = new EmailStoreFolder(FolderId, null, "Apple Mail");
        store.MutableFolders.Add(folder);

        EmailDocument document = emailResult.Document;
        bool isPartial = itemName.EndsWith(".partial.emlx", StringComparison.OrdinalIgnoreCase);
        document.Properties["EmailStore:Format"] = EmailStoreFormat.Emlx.ToString();
        document.Properties["EmailStore:ItemId"] = itemId;
        document.Properties["EmailStore:FolderId"] = FolderId;
        document.Properties["Emlx:SourceName"] = itemName;
        document.Properties["Emlx:DeclaredMessageBytes"] = declaredMessageBytes;
        document.Properties["Emlx:IsPartial"] = isPartial;

        ReadMetadata(stream, document, itemName, cancellationToken);
        if (isPartial) {
            _diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_EMLX_PARTIAL_MESSAGE",
                "This partial EMLX contains only the locally available RFC message; Apple Mail may keep additional body or attachment content in sibling storage.",
                EmailStoreDiagnosticSeverity.Warning,
                itemName));
        }

        EmailStoreItemReadParts loadedParts = EmailStoreItemReadParts.All;
        if (!(_includeAttachmentContent ?? _options.RetainAttachmentContent))
            loadedParts &= ~EmailStoreItemReadParts.AttachmentContent;
        folder.MutableItems.Add(new EmailStoreItem(
            itemId, FolderId, document, loadedParts: loadedParts, format: EmailStoreFormat.Emlx));
        return new EmailStoreReadResult(store, _diagnostics.AsReadOnly(), stream.Length);
    }

    private EmailReadResult ReadMessage(byte[] messageBytes, CancellationToken cancellationToken) {
        return EmailStoreMessageReader.Read(messageBytes, _options, cancellationToken,
            _includeAttachmentContent);
    }

    private void ReadMetadata(Stream stream, EmailDocument document, string itemName,
        CancellationToken cancellationToken) {
        long metadataLength = stream.Length - stream.Position;
        if (metadataLength == 0) return;
        if (metadataLength > _options.MaxDecodedPropertyBytesPerItem) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxDecodedPropertyBytesPerItem),
                metadataLength, _options.MaxDecodedPropertyBytesPerItem);
        }
        if (metadataLength > int.MaxValue) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxDecodedPropertyBytesPerItem),
                metadataLength, Math.Min(_options.MaxDecodedPropertyBytesPerItem, int.MaxValue));
        }

        byte[] metadata = ReadExact(stream, (int)metadataLength, cancellationToken,
            "The EMLX metadata trailer ended unexpectedly.");
        int contentOffset = FirstContentOffset(metadata);
        if (contentOffset == metadata.Length) return;
        try {
            if (EmlxPlistReader.LooksLikeBinaryPlist(metadata, contentOffset)) {
                _diagnostics.Add(new EmailStoreDiagnostic(
                    "EMAIL_STORE_EMLX_METADATA_UNSUPPORTED",
                    "The binary property-list trailer was not decoded; the RFC message remains available.",
                    EmailStoreDiagnosticSeverity.Warning,
                    itemName));
                return;
            }
            IReadOnlyDictionary<string, object?> values = EmlxPlistReader.Read(
                metadata, contentOffset, _options, cancellationToken);
            ApplyMetadata(document, values);
        } catch (EmailStoreLimitExceededException) {
            throw;
        } catch (Exception exception) when (exception is InvalidDataException || exception is FormatException ||
                                             exception is System.Xml.XmlException) {
            _diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_EMLX_METADATA_INVALID",
                exception.Message,
                EmailStoreDiagnosticSeverity.Warning,
                itemName));
        }
    }

    private static void ApplyMetadata(EmailDocument document, IReadOnlyDictionary<string, object?> values) {
        foreach (KeyValuePair<string, object?> pair in values) {
            document.Properties[string.Concat("Emlx:Metadata:", pair.Key)] = pair.Value;
        }

        if (TryInt64(values, "flags", out long flags)) ApplyFlags(document, flags);
        if (TryInt64(values, "date-received", out long received)) {
            document.ReceivedDate = FromUnixSeconds(received);
        }
        if (!document.Date.HasValue && TryInt64(values, "date-sent", out long sent)) {
            document.Date = FromUnixSeconds(sent);
        }
        if (document.Subject == null && values.TryGetValue("subject", out object? subject) && subject is string text) {
            document.Subject = text;
        }
        if (document.MessageId == null && values.TryGetValue("message-id", out object? messageId) &&
            messageId is string identifier) document.MessageId = identifier;
    }

    private static void ApplyFlags(EmailDocument document, long flags) {
        document.MessageMetadata.IsRead = (flags & (1L << 0)) != 0;
        document.MessageMetadata.IsDraft = document.MessageMetadata.IsDraft || (flags & (1L << 6)) != 0;
        document.Properties["Emlx:Flag:Deleted"] = (flags & (1L << 1)) != 0;
        document.Properties["Emlx:Flag:Answered"] = (flags & (1L << 2)) != 0;
        document.Properties["Emlx:Flag:Encrypted"] = (flags & (1L << 3)) != 0;
        document.Properties["Emlx:Flag:Flagged"] = (flags & (1L << 4)) != 0;
        document.Properties["Emlx:Flag:Recent"] = (flags & (1L << 5)) != 0;
        document.Properties["Emlx:Flag:Initial"] = (flags & (1L << 7)) != 0;
        document.Properties["Emlx:Flag:Forwarded"] = (flags & (1L << 8)) != 0;
        document.Properties["Emlx:Flag:Redirected"] = (flags & (1L << 9)) != 0;
        document.Properties["Emlx:Flag:AttachmentCount"] = (int)((flags >> 10) & 0x3F);
        document.Properties["Emlx:Flag:PriorityLevel"] = (int)((flags >> 16) & 0x7F);
        document.Properties["Emlx:Flag:Signed"] = (flags & (1L << 23)) != 0;
        document.Properties["Emlx:Flag:IsJunk"] = (flags & (1L << 24)) != 0;
        document.Properties["Emlx:Flag:IsNotJunk"] = (flags & (1L << 25)) != 0;
    }

    private void CopyDiagnostics(IEnumerable<EmailDiagnostic> diagnostics) {
        foreach (EmailDiagnostic diagnostic in diagnostics) {
            _diagnostics.Add(new EmailStoreDiagnostic(
                diagnostic.Code,
                diagnostic.Message,
                diagnostic.Severity == EmailDiagnosticSeverity.Error
                    ? EmailStoreDiagnosticSeverity.Error
                    : diagnostic.Severity == EmailDiagnosticSeverity.Information
                        ? EmailStoreDiagnosticSeverity.Information
                        : EmailStoreDiagnosticSeverity.Warning,
                diagnostic.Location));
        }
    }

    private static long ReadLengthPrefix(Stream stream, CancellationToken cancellationToken) {
        var prefix = new byte[MaxLengthPrefixBytes];
        int count = 0;
        while (count < prefix.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int value = stream.ReadByte();
            if (value < 0) throw new InvalidDataException("The EMLX length prefix is missing its line terminator.");
            if (value == '\n') break;
            prefix[count++] = (byte)value;
        }
        if (count == prefix.Length) throw new InvalidDataException("The EMLX length prefix is too long.");

        int start = 0;
        int end = count;
        while (start < end && (prefix[start] == ' ' || prefix[start] == '\t')) start++;
        while (end > start && (prefix[end - 1] == ' ' || prefix[end - 1] == '\t' || prefix[end - 1] == '\r')) end--;
        if (start == end) throw new InvalidDataException("The EMLX length prefix is empty.");
        long result = 0;
        for (int index = start; index < end; index++) {
            byte value = prefix[index];
            if (value < '0' || value > '9') throw new InvalidDataException("The EMLX length prefix must contain decimal digits only.");
            int digit = value - '0';
            if (result > (long.MaxValue - digit) / 10) throw new InvalidDataException("The EMLX length prefix exceeds Int64.");
            result = result * 10 + digit;
        }
        return result;
    }

    private static byte[] ReadExact(Stream stream, int length, CancellationToken cancellationToken, string error) {
        var data = new byte[length];
        int offset = 0;
        while (offset < data.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int read = stream.Read(data, offset, data.Length - offset);
            if (read == 0) throw new InvalidDataException(error);
            offset += read;
        }
        return data;
    }

    private static int FirstContentOffset(byte[] data) {
        int offset = 0;
        while (offset < data.Length && (data[offset] == ' ' || data[offset] == '\t' ||
                                        data[offset] == '\r' || data[offset] == '\n')) offset++;
        return offset;
    }

    private static bool TryInt64(IReadOnlyDictionary<string, object?> values, string key, out long value) {
        if (values.TryGetValue(key, out object? item)) {
            if (item is long integer) { value = integer; return true; }
            if (item is int smallInteger) { value = smallInteger; return true; }
        }
        value = 0;
        return false;
    }

    private static DateTimeOffset? FromUnixSeconds(long seconds) {
        try {
            return new DateTimeOffset(1970, 1, 1, 0, 0, 0, TimeSpan.Zero).AddSeconds(seconds);
        } catch (ArgumentOutOfRangeException) {
            return null;
        }
    }

    private static string GetItemName(string? sourceName) {
        if (sourceName == null || sourceName.Trim().Length == 0) return "message.emlx";
        try {
            return Path.GetFileName(sourceName);
        } catch (Exception exception) when (exception is ArgumentException || exception is NotSupportedException) {
            return sourceName ?? "message.emlx";
        }
    }

    private static string GetDisplayName(string itemName) {
        string name = itemName.EndsWith(".emlx", StringComparison.OrdinalIgnoreCase)
            ? itemName.Substring(0, itemName.Length - 5)
            : itemName;
        return name.EndsWith(".partial", StringComparison.OrdinalIgnoreCase)
            ? name.Substring(0, name.Length - 8)
            : name;
    }
}
