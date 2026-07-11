using OfficeIMO.Shared;

namespace OfficeIMO.Email;

/// <summary>Reads bounded email and Outlook artifacts into the shared <see cref="EmailDocument"/> model.</summary>
public sealed class EmailDocumentReader {
    private static readonly byte[] CompoundSignature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
    private static readonly byte[] TnefSignature = { 0x78, 0x9F, 0x3E, 0x22 };
    private readonly EmailReaderOptions _options;

    /// <summary>Creates a reader with the default bounded policy.</summary>
    public EmailDocumentReader() : this(EmailReaderOptions.Default) { }

    /// <summary>Creates a reader with an immutable bounded policy.</summary>
    public EmailDocumentReader(EmailReaderOptions options) {
        _options = options ?? throw new ArgumentNullException(nameof(options));
    }

    /// <summary>Reader policy used by this instance.</summary>
    public EmailReaderOptions Options => _options;

    /// <summary>Reads an artifact from a file.</summary>
    public EmailReadResult Read(string filePath, CancellationToken cancellationToken = default) {
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        using (FileStream stream = File.OpenRead(filePath)) return Read(stream, cancellationToken);
    }

    /// <summary>Reads an artifact from memory.</summary>
    public EmailReadResult Read(byte[] data, CancellationToken cancellationToken = default) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        cancellationToken.ThrowIfCancellationRequested();
        if (data.LongLength > _options.MaxInputBytes) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxInputBytes), data.LongLength, _options.MaxInputBytes);
        }
        return Parse(data, cancellationToken);
    }

    /// <summary>Reads an artifact from the stream's current position without closing it.</summary>
    public EmailReadResult Read(Stream stream, CancellationToken cancellationToken = default) {
        return Parse(EmailByteReader.ReadAll(stream, _options.MaxInputBytes, cancellationToken), cancellationToken);
    }

    /// <summary>Asynchronously reads an artifact from a file.</summary>
    public async Task<EmailReadResult> ReadAsync(string filePath, CancellationToken cancellationToken = default) {
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read,
            81920, FileOptions.Asynchronous | FileOptions.SequentialScan)) {
            return await ReadAsync(stream, cancellationToken).ConfigureAwait(false);
        }
    }

    /// <summary>Asynchronously reads an artifact from the stream's current position without closing it.</summary>
    public async Task<EmailReadResult> ReadAsync(Stream stream, CancellationToken cancellationToken = default) {
        byte[] data = await EmailByteReader.ReadAllAsync(stream, _options.MaxInputBytes, cancellationToken).ConfigureAwait(false);
        return Parse(data, cancellationToken);
    }

    /// <summary>Detects the artifact format from content rather than the filename.</summary>
    public static EmailFileFormat DetectFormat(byte[] data) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (StartsWith(data, CompoundSignature)) {
            return OfficeCompoundFileReader.TryRead(data, out OfficeCompoundFile? compound, out _) &&
                compound != null && compound.Streams.ContainsKey("__properties_version1.0")
                ? EmailFileFormat.OutlookMsg
                : EmailFileFormat.Unknown;
        }
        if (StartsWith(data, TnefSignature)) return EmailFileFormat.Tnef;
        if (data.Length >= 5 && data[0] == 'F' && data[1] == 'r' && data[2] == 'o' && data[3] == 'm' && data[4] == ' ') {
            return EmailFileFormat.Mbox;
        }
        return LooksLikeMessage(data) ? EmailFileFormat.Eml : EmailFileFormat.Unknown;
    }

    private EmailReadResult Parse(byte[] data, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        List<EmailDiagnostic> diagnostics = new List<EmailDiagnostic>();
        EmailDocument document;
        if (StartsWith(data, CompoundSignature)) {
            if (!MsgReader.TryRead(data, _options, diagnostics, cancellationToken, out document)) {
                diagnostics.Add(new EmailDiagnostic("EMAIL_FORMAT_UNKNOWN",
                    "The compound artifact is not an Outlook MSG item.", EmailDiagnosticSeverity.Error));
            }
            cancellationToken.ThrowIfCancellationRequested();
            if (_options.PreserveRawSource) document.RawSource = (byte[])data.Clone();
            return new EmailReadResult(document, diagnostics.AsReadOnly(), data.LongLength);
        }

        EmailFileFormat format = DetectFormat(data);
        switch (format) {
            case EmailFileFormat.Eml:
                document = MimeParser.Parse(data, _options, diagnostics, cancellationToken);
                break;
            case EmailFileFormat.Tnef:
                document = TnefReader.Read(data, _options, diagnostics, cancellationToken);
                break;
            case EmailFileFormat.Unknown:
                diagnostics.Add(new EmailDiagnostic("EMAIL_FORMAT_UNKNOWN",
                    "The artifact has no recognized email signature or RFC message header.", EmailDiagnosticSeverity.Error));
                document = new EmailDocument { Format = EmailFileFormat.Unknown, OutlookItemKind = OutlookItemKind.Unknown };
                break;
            case EmailFileFormat.Mbox:
                diagnostics.Add(new EmailDiagnostic("EMAIL_MBOX_REQUIRES_MAILBOX_READER",
                    "Use EmailMailboxReader to read all messages from an mbox aggregate.", EmailDiagnosticSeverity.Error));
                document = new EmailDocument { Format = EmailFileFormat.Mbox, OutlookItemKind = OutlookItemKind.Unknown };
                break;
            case EmailFileFormat.OutlookMsg:
                throw new InvalidOperationException("MSG input must be handled by the compound-file read path.");
            default:
                diagnostics.Add(new EmailDiagnostic("EMAIL_FORMAT_NOT_IMPLEMENTED",
                    string.Concat(format.ToString(), " support is not available in this delivery slice."), EmailDiagnosticSeverity.Error));
                document = new EmailDocument { Format = format, OutlookItemKind = OutlookItemKind.Unknown };
                break;
        }
        cancellationToken.ThrowIfCancellationRequested();
        if (_options.PreserveRawSource) document.RawSource = (byte[])data.Clone();
        return new EmailReadResult(document, diagnostics.AsReadOnly(), data.LongLength);
    }

    private static bool StartsWith(byte[] data, byte[] signature) {
        if (data.Length < signature.Length) return false;
        for (int i = 0; i < signature.Length; i++) {
            if (data[i] != signature[i]) return false;
        }
        return true;
    }

    private static bool LooksLikeMessage(byte[] data) {
        int limit = Math.Min(data.Length, 64 * 1024);
        int lineStart = data.Length >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF ? 3 : 0;
        while (lineStart < limit) {
            int lineEnd = lineStart;
            while (lineEnd < limit && data[lineEnd] != '\r' && data[lineEnd] != '\n') lineEnd++;
            if (lineEnd == lineStart) return false;
            for (int i = lineStart; i < lineEnd; i++) {
                if (data[i] == ':' && i > lineStart) return true;
                if (!IsHeaderNameCharacter(data[i])) break;
            }
            lineStart = lineEnd;
            if (lineStart < limit && data[lineStart] == '\r') lineStart++;
            if (lineStart < limit && data[lineStart] == '\n') lineStart++;
        }
        return false;
    }

    private static bool IsHeaderNameCharacter(byte value) {
        return value >= 'a' && value <= 'z' || value >= 'A' && value <= 'Z' ||
            value >= '0' && value <= '9' || value == '!' || value == '#' || value == '$' ||
            value == '%' || value == '&' || value == '\'' || value == '*' || value == '+' ||
            value == '-' || value == '.' || value == '^' || value == '_' || value == '`' ||
            value == '|' || value == '~';
    }
}
