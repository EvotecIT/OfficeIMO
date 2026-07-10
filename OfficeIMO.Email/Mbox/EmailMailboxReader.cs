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
        if (data.LongLength > _options.MessageOptions.MaxInputBytes) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxInputBytes), data.LongLength,
                _options.MessageOptions.MaxInputBytes);
        }
        return Parse(data, cancellationToken);
    }

    /// <summary>Reads from the stream's current position without closing it.</summary>
    public EmailMailboxReadResult Read(Stream stream, CancellationToken cancellationToken = default) {
        return Parse(EmailByteReader.ReadAll(stream, _options.MessageOptions.MaxInputBytes, cancellationToken), cancellationToken);
    }

    /// <summary>Asynchronously reads a mailbox file.</summary>
    public async Task<EmailMailboxReadResult> ReadAsync(string filePath, CancellationToken cancellationToken = default) {
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read,
            81920, FileOptions.Asynchronous | FileOptions.SequentialScan)) {
            return await ReadAsync(stream, cancellationToken).ConfigureAwait(false);
        }
    }

    /// <summary>Asynchronously reads from the stream's current position without closing it.</summary>
    public async Task<EmailMailboxReadResult> ReadAsync(Stream stream, CancellationToken cancellationToken = default) {
        byte[] data = await EmailByteReader.ReadAllAsync(stream, _options.MessageOptions.MaxInputBytes, cancellationToken).ConfigureAwait(false);
        return Parse(data, cancellationToken);
    }

    private EmailMailboxReadResult Parse(byte[] data, CancellationToken cancellationToken) {
        var diagnostics = new List<EmailDiagnostic>();
        var mailbox = new EmailMailbox();
        List<Envelope> envelopes = FindEnvelopes(data, cancellationToken);
        if (envelopes.Count == 0 || envelopes[0].LineStart != 0) {
            diagnostics.Add(new EmailDiagnostic("EMAIL_MBOX_ENVELOPE_MISSING",
                "The mailbox does not begin with an mbox From separator.", EmailDiagnosticSeverity.Error));
            return new EmailMailboxReadResult(mailbox, diagnostics.AsReadOnly(), data.LongLength);
        }
        if (envelopes.Count > _options.MaxMessageCount) {
            throw new EmailLimitExceededException(nameof(EmailMailboxReaderOptions.MaxMessageCount),
                envelopes.Count, _options.MaxMessageCount);
        }

        var reader = new EmailDocumentReader(_options.MessageOptions);
        for (int index = 0; index < envelopes.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            Envelope envelope = envelopes[index];
            int end = index + 1 < envelopes.Count ? envelopes[index + 1].LineStart : data.Length;
            byte[] messageBytes = MsgBinary.Slice(data, envelope.MessageStart, end - envelope.MessageStart);
            MboxVariant variant = _options.Variant == MboxVariant.Auto ? DetectVariant(messageBytes, cancellationToken) : _options.Variant;
            messageBytes = Unescape(messageBytes, variant, cancellationToken);
            EmailReadResult message = reader.Read(messageBytes, cancellationToken);
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

    private static List<Envelope> FindEnvelopes(byte[] data, CancellationToken cancellationToken) {
        var result = new List<Envelope>();
        int lineStart = 0;
        while (lineStart < data.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int lineEnd = lineStart;
            while (lineEnd < data.Length && data[lineEnd] != '\r' && data[lineEnd] != '\n') lineEnd++;
            if (lineEnd - lineStart >= 5 && StartsWith(data, lineStart, "From ")) {
                string raw = Encoding.ASCII.GetString(data, lineStart, lineEnd - lineStart);
                ParseEnvelope(raw, out string? sender, out DateTimeOffset? date);
                result.Add(new Envelope(lineStart, SkipLineEnding(data, lineEnd), raw, sender, date));
            }
            lineStart = SkipLineEnding(data, lineEnd);
        }
        return result;
    }

    private static void ParseEnvelope(string raw, out string? sender, out DateTimeOffset? date) {
        string remainder = raw.Substring(5).Trim();
        int separator = remainder.IndexOf(' ');
        sender = separator < 0 ? remainder : remainder.Substring(0, separator);
        string dateText = separator < 0 ? string.Empty : remainder.Substring(separator + 1).Trim();
        date = DateTimeOffset.TryParse(dateText, CultureInfo.InvariantCulture,
            DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeUniversal, out DateTimeOffset parsed) ? parsed : (DateTimeOffset?)null;
    }

    private static MboxVariant DetectVariant(byte[] message, CancellationToken cancellationToken) {
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

    private static byte[] Unescape(byte[] message, MboxVariant variant, CancellationToken cancellationToken) {
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
