namespace OfficeIMO.Email;

internal static class MboxStreamReader {
    internal static IEnumerable<EmailMailboxEntryReadResult> Read(Stream stream, EmailMailboxReaderOptions options,
        CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("The stream must be readable.", nameof(stream));
        return Enumerate(stream, options, cancellationToken);
    }

    private static IEnumerable<EmailMailboxEntryReadResult> Enumerate(Stream stream,
        EmailMailboxReaderOptions options, CancellationToken cancellationToken) {
        long originalPosition = stream.CanSeek ? stream.Position : 0;
        if (stream.CanSeek) stream.Position = 0;
        long totalBytes = 0;
        int messageCount = 0;
        Envelope? envelope = null;
        var message = new MemoryStream();
        long entryBytes = 0;
        int pendingByte = -1;
        try {
            while (TryReadLine(stream, cancellationToken, ref pendingByte, totalBytes, message.Length,
                       envelope != null, options, out byte[] line, out byte[] ending)) {
                long lineBytes = checked(line.LongLength + ending.LongLength);
                totalBytes = checked(totalBytes + lineBytes);
                if (totalBytes > options.MaxMailboxBytes) {
                    throw new EmailLimitExceededException(nameof(EmailMailboxReaderOptions.MaxMailboxBytes),
                        totalBytes, options.MaxMailboxBytes);
                }

                byte[] envelopeLine = envelope == null && messageCount == 0 && totalBytes == lineBytes && HasUtf8Bom(line)
                    ? Slice(line, 3)
                    : line;
                if (StartsWith(envelopeLine, "From ")) {
                    if (envelope != null) {
                        yield return ParseEntry(message.ToArray(), envelope, entryBytes, options, cancellationToken);
                        message.Dispose();
                        message = new MemoryStream();
                    } else if (totalBytes != lineBytes) {
                        throw new InvalidDataException(
                            "EMAIL_MBOX_ENVELOPE_MISSING: The mailbox does not begin with an mbox From separator.");
                    }

                    messageCount++;
                    if (messageCount > options.MaxMessageCount) {
                        throw new EmailLimitExceededException(nameof(EmailMailboxReaderOptions.MaxMessageCount),
                            messageCount, options.MaxMessageCount);
                    }
                    string rawLine = Encoding.ASCII.GetString(envelopeLine);
                    EmailMailboxReader.ParseEnvelope(rawLine, out string? sender, out DateTimeOffset? date);
                    envelope = new Envelope(rawLine, sender, date);
                    entryBytes = lineBytes;
                    continue;
                }

                if (envelope == null) {
                    throw new InvalidDataException(
                        "EMAIL_MBOX_ENVELOPE_MISSING: The mailbox does not begin with an mbox From separator.");
                }
                long messageLength = checked(message.Length + lineBytes);
                if (messageLength > options.MessageOptions.MaxInputBytes) {
                    throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxInputBytes), messageLength,
                        options.MessageOptions.MaxInputBytes);
                }
                message.Write(line, 0, line.Length);
                message.Write(ending, 0, ending.Length);
                entryBytes = checked(entryBytes + lineBytes);
            }

            if (envelope == null) {
                throw new InvalidDataException(
                    "EMAIL_MBOX_ENVELOPE_MISSING: The mailbox does not begin with an mbox From separator.");
            }
            yield return ParseEntry(message.ToArray(), envelope, entryBytes, options, cancellationToken);
        } finally {
            message.Dispose();
            if (stream.CanSeek) stream.Position = originalPosition;
        }
    }

    private static EmailMailboxEntryReadResult ParseEntry(byte[] messageBytes, Envelope envelope, long bytesRead,
        EmailMailboxReaderOptions options, CancellationToken cancellationToken) {
        MboxVariant variant = options.Variant == MboxVariant.Auto
            ? EmailMailboxReader.DetectVariant(messageBytes, cancellationToken)
            : options.Variant;
        byte[] unescaped = EmailMailboxReader.Unescape(messageBytes, variant, cancellationToken);
        var reader = new EmailDocumentReader(options.MessageOptions);
        EmailReadResult message = EmailMailboxReader.ReadEntryMessage(
            reader, unescaped, options.MessageOptions, cancellationToken);
        var entry = new EmailMailboxEntry(message.Document) {
            EnvelopeSender = envelope.Sender,
            EnvelopeDate = envelope.Date,
            RawFromLine = envelope.RawLine
        };
        return new EmailMailboxEntryReadResult(entry, message.Diagnostics, bytesRead);
    }

    private static bool TryReadLine(Stream stream, CancellationToken cancellationToken, ref int pendingByte,
        long totalBytes, long messageBytes, bool hasEnvelope, EmailMailboxReaderOptions options,
        out byte[] line, out byte[] ending) {
        using (var content = new MemoryStream()) {
            int current = pendingByte;
            pendingByte = -1;
            bool couldBeEnvelope = hasEnvelope;
            while (current >= 0 || (current = stream.ReadByte()) >= 0) {
                cancellationToken.ThrowIfCancellationRequested();
                if (current == '\n') {
                    bool envelopeCandidate = couldBeEnvelope && content.Length >= 5;
                    EnsureLineWithinLimits(content.Length + 1, envelopeCandidate, totalBytes, messageBytes,
                        hasEnvelope, options);
                    line = content.ToArray();
                    if (line.Length > 0 && line[line.Length - 1] == '\r') {
                        Array.Resize(ref line, line.Length - 1);
                        ending = new byte[] { (byte)'\r', (byte)'\n' };
                    } else ending = new byte[] { (byte)'\n' };
                    return true;
                }
                if (current == '\r') {
                    int next = stream.ReadByte();
                    if (next == '\n') {
                        bool envelopeCandidate = couldBeEnvelope && content.Length >= 5;
                        EnsureLineWithinLimits(content.Length + 2, envelopeCandidate, totalBytes, messageBytes,
                            hasEnvelope, options);
                        line = content.ToArray();
                        ending = new byte[] { (byte)'\r', (byte)'\n' };
                        return true;
                    }
                    bool bareEnvelopeCandidate = couldBeEnvelope && content.Length >= 5;
                    EnsureLineWithinLimits(content.Length + 1, bareEnvelopeCandidate, totalBytes, messageBytes,
                        hasEnvelope, options);
                    line = content.ToArray();
                    ending = new byte[] { (byte)'\r' };
                    pendingByte = next;
                    return true;
                }
                if (couldBeEnvelope && content.Length < 5 &&
                    current != "From "[(int)content.Length]) couldBeEnvelope = false;
                EnsureLineWithinLimits(content.Length + 1, couldBeEnvelope, totalBytes, messageBytes,
                    hasEnvelope, options);
                content.WriteByte((byte)current);
                current = -1;
            }
            if (content.Length == 0) {
                line = Array.Empty<byte>();
                ending = Array.Empty<byte>();
                return false;
            }
            EnsureLineWithinLimits(content.Length, couldBeEnvelope && content.Length >= 5, totalBytes,
                messageBytes, hasEnvelope, options);
            line = content.ToArray();
            ending = Array.Empty<byte>();
            return true;
        }
    }

    private static void EnsureLineWithinLimits(long lineBytes, bool couldBeEnvelope, long totalBytes,
        long messageBytes, bool hasEnvelope, EmailMailboxReaderOptions options) {
        long aggregateLength = checked(totalBytes + lineBytes);
        if (aggregateLength > options.MaxMailboxBytes) {
            throw new EmailLimitExceededException(nameof(EmailMailboxReaderOptions.MaxMailboxBytes),
                aggregateLength, options.MaxMailboxBytes);
        }
        if (!hasEnvelope || couldBeEnvelope) return;
        long messageLength = checked(messageBytes + lineBytes);
        if (messageLength > options.MessageOptions.MaxInputBytes) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxInputBytes), messageLength,
                options.MessageOptions.MaxInputBytes);
        }
    }

    private static bool StartsWith(byte[] data, string value) {
        if (data.Length < value.Length) return false;
        for (int index = 0; index < value.Length; index++) if (data[index] != value[index]) return false;
        return true;
    }

    private static bool HasUtf8Bom(byte[] data) => data.Length >= 3 &&
        data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF;

    private static byte[] Slice(byte[] data, int offset) {
        var result = new byte[data.Length - offset];
        Buffer.BlockCopy(data, offset, result, 0, result.Length);
        return result;
    }

    private sealed class Envelope {
        internal Envelope(string rawLine, string? sender, DateTimeOffset? date) {
            RawLine = rawLine; Sender = sender; Date = date;
        }
        internal string RawLine { get; }
        internal string? Sender { get; }
        internal DateTimeOffset? Date { get; }
    }
}
