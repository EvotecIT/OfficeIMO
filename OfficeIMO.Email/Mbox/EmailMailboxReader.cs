namespace OfficeIMO.Email;

/// <summary>Reads mboxo and mboxrd mailbox aggregates.</summary>
public sealed class EmailMailboxReader {
    private readonly EmailMailboxReaderOptions _options;

    /// <summary>Creates a reader with the default policy.</summary>
    public EmailMailboxReader() : this(EmailMailboxReaderOptions.Default) { }

    /// <summary>Creates a reader with an immutable policy.</summary>
    public EmailMailboxReader(EmailMailboxReaderOptions options) {
        _options = options ?? throw new ArgumentNullException(nameof(options));
    }

    /// <summary>Reader policy used by this instance.</summary>
    public EmailMailboxReaderOptions Options => _options;

    /// <summary>Reads a mailbox file.</summary>
    public EmailMailboxReadResult Read(string filePath, CancellationToken cancellationToken = default) {
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        using (FileStream stream = File.OpenRead(filePath)) return Read(stream, cancellationToken);
    }

    /// <summary>Reads mailbox bytes.</summary>
    public EmailMailboxReadResult Read(byte[] data, CancellationToken cancellationToken = default) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        cancellationToken.ThrowIfCancellationRequested();
        if (data.LongLength > _options.MaxMailboxBytes) {
            throw new EmailLimitExceededException(nameof(EmailMailboxReaderOptions.MaxMailboxBytes), data.LongLength,
                _options.MaxMailboxBytes);
        }
        return Parse(data, cancellationToken);
    }

    /// <summary>
    /// Reads without closing the stream. Seekable streams are read from the beginning and restored to their
    /// original position; non-seekable streams are read forward from their current position.
    /// </summary>
    public EmailMailboxReadResult Read(Stream stream, CancellationToken cancellationToken = default) {
        byte[] data;
        try {
            data = EmailByteReader.ReadAll(stream, _options.MaxMailboxBytes, cancellationToken);
        } catch (EmailLimitExceededException exception) when
            (exception.LimitName == nameof(EmailReaderOptions.MaxInputBytes)) {
            throw CreateMailboxLimitException(exception);
        }
        return Parse(data, cancellationToken);
    }

    /// <summary>Asynchronously reads a mailbox file.</summary>
    public async Task<EmailMailboxReadResult> ReadAsync(string filePath, CancellationToken cancellationToken = default) {
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read,
            81920, FileOptions.Asynchronous | FileOptions.SequentialScan)) {
            return await ReadAsync(stream, cancellationToken).ConfigureAwait(false);
        }
    }

    /// <summary>
    /// Asynchronously reads without closing the stream. Seekable streams are read from the beginning and restored
    /// to their original position; non-seekable streams are read forward.
    /// </summary>
    public async Task<EmailMailboxReadResult> ReadAsync(Stream stream, CancellationToken cancellationToken = default) {
        byte[] data;
        try {
            data = await EmailByteReader.ReadAllAsync(stream, _options.MaxMailboxBytes, cancellationToken)
                .ConfigureAwait(false);
        } catch (EmailLimitExceededException exception) when
            (exception.LimitName == nameof(EmailReaderOptions.MaxInputBytes)) {
            throw CreateMailboxLimitException(exception);
        }
        return Parse(data, cancellationToken);
    }

    /// <summary>Enumerates a mailbox file while retaining at most one decoded message in memory.</summary>
    public IEnumerable<EmailMailboxEntryReadResult> ReadEntries(string filePath,
        CancellationToken cancellationToken = default) {
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        return EnumerateFile(filePath, cancellationToken);
    }

    /// <summary>
    /// Enumerates a caller-owned mailbox stream while retaining at most one decoded message in memory.
    /// Seekable streams are restored to their original position when enumeration ends or is disposed.
    /// </summary>
    public IEnumerable<EmailMailboxEntryReadResult> ReadEntries(Stream stream,
        CancellationToken cancellationToken = default) => MboxStreamReader.Read(stream, _options, cancellationToken);

    private IEnumerable<EmailMailboxEntryReadResult> EnumerateFile(string filePath,
        CancellationToken cancellationToken) {
        using (FileStream stream = File.OpenRead(filePath)) {
            foreach (EmailMailboxEntryReadResult entry in ReadEntries(stream, cancellationToken)) yield return entry;
        }
    }

    private static EmailLimitExceededException CreateMailboxLimitException(EmailLimitExceededException exception) =>
        new EmailLimitExceededException(nameof(EmailMailboxReaderOptions.MaxMailboxBytes),
            exception.ActualValue, exception.MaximumValue);

    private EmailMailboxReadResult Parse(byte[] data, CancellationToken cancellationToken) {
        var diagnostics = new List<EmailDiagnostic>();
        var mailbox = new EmailMailbox();
        List<Envelope> envelopes = FindEnvelopes(data, _options.MaxMessageCount, cancellationToken);
        if (envelopes.Count == 0 || envelopes[0].LineStart != 0) {
            diagnostics.Add(new EmailDiagnostic("EMAIL_MBOX_ENVELOPE_MISSING",
                "The mailbox does not begin with an mbox From separator.", EmailDiagnosticSeverity.Error));
            return new EmailMailboxReadResult(mailbox, diagnostics.AsReadOnly(), data.LongLength);
        }
        var reader = new EmailDocumentReader(_options.MessageOptions);
        for (int index = 0; index < envelopes.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            Envelope envelope = envelopes[index];
            int end = index + 1 < envelopes.Count ? envelopes[index + 1].LineStart : data.Length;
            byte[] messageBytes = MsgBinary.Slice(data, envelope.MessageStart, end - envelope.MessageStart);
            MboxVariant variant = _options.Variant == MboxVariant.Auto ? DetectVariant(messageBytes, cancellationToken) : _options.Variant;
            messageBytes = Unescape(messageBytes, variant, cancellationToken);
            EmailReadResult message = ReadEntryMessage(reader, messageBytes, _options.MessageOptions, cancellationToken);
            var entry = new EmailMailboxEntry(message.Document) {
                EnvelopeSender = envelope.Sender,
                EnvelopeDate = envelope.Date,
                RawFromLine = envelope.RawLine
            };
            mailbox.Messages.Add(entry);
            foreach (EmailDiagnostic diagnostic in message.Diagnostics) {
                diagnostics.Add(new EmailDiagnostic(diagnostic.Code, diagnostic.Message, diagnostic.Severity,
                    string.Concat("message[", index.ToString(CultureInfo.InvariantCulture), "]",
                        diagnostic.Location == null ? string.Empty : string.Concat("/", diagnostic.Location))));
            }
        }
        return new EmailMailboxReadResult(mailbox, diagnostics.AsReadOnly(), data.LongLength);
    }

    private static List<Envelope> FindEnvelopes(byte[] data, int maximumMessages,
        CancellationToken cancellationToken) {
        var result = new List<Envelope>();
        int lineStart = HasUtf8Bom(data) ? 3 : 0;
        bool firstLine = true;
        while (lineStart < data.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int lineEnd = lineStart;
            while (lineEnd < data.Length && data[lineEnd] != '\r' && data[lineEnd] != '\n') lineEnd++;
            if (lineEnd - lineStart >= 5 && StartsWith(data, lineStart, "From ")) {
                string raw = Encoding.ASCII.GetString(data, lineStart, lineEnd - lineStart);
                ParseEnvelope(raw, out string? sender, out DateTimeOffset? date);
                result.Add(new Envelope(firstLine && lineStart == 3 ? 0 : lineStart,
                    SkipLineEnding(data, lineEnd), raw, sender, date));
                if (result.Count > maximumMessages) {
                    throw new EmailLimitExceededException(nameof(EmailMailboxReaderOptions.MaxMessageCount),
                        result.Count, maximumMessages);
                }
            }
            lineStart = SkipLineEnding(data, lineEnd);
            firstLine = false;
        }
        return result;
    }

    internal static void ParseEnvelope(string raw, out string? sender, out DateTimeOffset? date) {
        string remainder = raw.Substring(5).Trim();
        int separator = remainder.IndexOf(' ');
        sender = separator < 0 ? remainder : remainder.Substring(0, separator);
        string dateText = separator < 0 ? string.Empty : remainder.Substring(separator + 1).Trim();
        const DateTimeStyles dateStyles = DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeUniversal;
        DateTimeOffset parsed;
        if (DateTimeOffset.TryParseExact(dateText, "ddd MMM dd HH:mm:ss yyyy",
                CultureInfo.InvariantCulture, dateStyles, out parsed) ||
            DateTimeOffset.TryParse(dateText, CultureInfo.InvariantCulture, dateStyles, out parsed)) {
            date = parsed;
        } else {
            date = null;
        }
    }

    internal static MboxVariant DetectVariant(byte[] message, CancellationToken cancellationToken) {
        int lineStart = 0;
        while (lineStart < message.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int lineEnd = lineStart;
            while (lineEnd < message.Length && message[lineEnd] != '\r' && message[lineEnd] != '\n') lineEnd++;
            int position = lineStart;
            int angles = 0;
            while (position < lineEnd && message[position] == '>') { angles++; position++; }
            if (angles > 1 && StartsWith(message, position, "From ")) return MboxVariant.Mboxrd;
            lineStart = SkipLineEnding(message, lineEnd);
        }
        return MboxVariant.Mboxo;
    }

    internal static byte[] Unescape(byte[] message, MboxVariant variant, CancellationToken cancellationToken) {
        using (MemoryStream output = new MemoryStream(message.Length)) {
            int lineStart = 0;
            while (lineStart < message.Length) {
                cancellationToken.ThrowIfCancellationRequested();
                int lineEnd = lineStart;
                while (lineEnd < message.Length && message[lineEnd] != '\r' && message[lineEnd] != '\n') lineEnd++;
                int writeStart = lineStart;
                if (message[lineStart] == '>') {
                    int position = lineStart + 1;
                    if (variant == MboxVariant.Mboxrd) while (position < lineEnd && message[position] == '>') position++;
                    if (StartsWith(message, position, "From ")) writeStart++;
                }
                output.Write(message, writeStart, lineEnd - writeStart);
                int after = SkipLineEnding(message, lineEnd);
                output.Write(message, lineEnd, after - lineEnd);
                lineStart = after;
            }
            return output.ToArray();
        }
    }

    private static bool StartsWith(byte[] data, int offset, string value) {
        if (offset < 0 || offset + value.Length > data.Length) return false;
        for (int index = 0; index < value.Length; index++) if (data[offset + index] != value[index]) return false;
        return true;
    }

    internal static EmailReadResult ReadEntryMessage(EmailDocumentReader reader, byte[] messageBytes,
        EmailReaderOptions options, CancellationToken cancellationToken) {
        EmailReadResult result = reader.Read(messageBytes, cancellationToken);
        if (result.Document.Format != EmailFileFormat.Unknown ||
            !result.Diagnostics.Any(diagnostic => diagnostic.Code == "EMAIL_FORMAT_UNKNOWN")) return result;

        var diagnostics = result.Diagnostics
            .Where(diagnostic => diagnostic.Code != "EMAIL_FORMAT_UNKNOWN")
            .ToList();
        diagnostics.Add(new EmailDiagnostic("EMAIL_MBOX_MESSAGE_HEADERS_MISSING",
            "An mbox entry without recognizable message headers was retained as a plain MIME body.",
            EmailDiagnosticSeverity.Warning));
        byte[] mimeBody = new byte[messageBytes.Length + 2];
        mimeBody[0] = (byte)'\r';
        mimeBody[1] = (byte)'\n';
        Buffer.BlockCopy(messageBytes, 0, mimeBody, 2, messageBytes.Length);
        EmailDocument document = MimeParser.Parse(mimeBody, options, diagnostics, cancellationToken);
        return new EmailReadResult(document, diagnostics.AsReadOnly(), messageBytes.LongLength);
    }

    private static bool HasUtf8Bom(byte[] data) => data.Length >= 3 &&
        data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF;

    private static int SkipLineEnding(byte[] data, int position) {
        if (position < data.Length && data[position] == '\r') position++;
        if (position < data.Length && data[position] == '\n') position++;
        return position;
    }

    private sealed class Envelope {
        internal Envelope(int lineStart, int messageStart, string rawLine, string? sender, DateTimeOffset? date) {
            LineStart = lineStart; MessageStart = messageStart; RawLine = rawLine; Sender = sender; Date = date;
        }
        internal int LineStart { get; }
        internal int MessageStart { get; }
        internal string RawLine { get; }
        internal string? Sender { get; }
        internal DateTimeOffset? Date { get; }
    }
}
