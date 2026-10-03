using OfficeIMO.Core.Internal;

namespace OfficeIMO.Email;

/// <summary>Spools one indexed mbox entry without materializing its encoded MIME or decoded attachments.</summary>
internal static class MboxSelectedMessageReader {
    internal static EmailReadResult Read(Stream input, EmailMailboxReaderOptions options,
        CancellationToken cancellationToken, out EmailMailboxEntry entry) {
        if (input.Length > options.MaxMailboxBytes)
            throw new EmailLimitExceededException(nameof(EmailMailboxReaderOptions.MaxMailboxBytes), input.Length, options.MaxMailboxBytes);
        using var buffered = new BufferedStream(input, 64 * 1024);
        buffered.Position = 0;
        string envelope = ReadEnvelope(buffered, options.MessageOptions.MaxHeaderBytes, cancellationToken);
        long start = buffered.Position;
        long length = buffered.Length - start;
        if (length > options.MessageOptions.MaxInputBytes)
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxInputBytes), length, options.MessageOptions.MaxInputBytes);
        MboxVariant variant = options.Variant == MboxVariant.Auto ? DetectVariant(buffered, cancellationToken) : options.Variant;
        buffered.Position = start;
        using FileStream spool = OfficeTemporaryFile.Create("OfficeIMO.Email.Mbox.", ".eml", FileOptions.SequentialScan, out _);
        using (var output = new BufferedStream(spool, 64 * 1024)) {
            Unescape(buffered, output, variant, cancellationToken);
            output.Flush();
            output.Position = 0;
            EmailReadResult result = EmailMailboxReader.ReadEntryMessage(new EmailDocumentReader(options.MessageOptions),
                output, options.MessageOptions, cancellationToken);
            EmailMailboxReader.ParseEnvelope(envelope, out string? sender, out DateTimeOffset? date);
            entry = new EmailMailboxEntry(result.Document) { EnvelopeSender = sender, EnvelopeDate = date, RawFromLine = envelope };
            return result;
        }
    }

    private static string ReadEnvelope(Stream input, int maximum, CancellationToken cancellationToken) {
        using var line = new MemoryStream();
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            int value = input.ReadByte();
            if (value < 0 || value == '\n') break;
            if (value == '\r') {
                int next = input.ReadByte();
                if (next >= 0 && next != '\n') input.Position--;
                break;
            }
            if (line.Length >= maximum)
                throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxHeaderBytes), line.Length + 1, maximum);
            line.WriteByte((byte)value);
        }
        byte[] bytes = line.ToArray();
        int offset = bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF ? 3 : 0;
        string raw = Encoding.ASCII.GetString(bytes, offset, bytes.Length - offset);
        if (!raw.StartsWith("From ", StringComparison.Ordinal))
            throw new InvalidDataException("EMAIL_MBOX_ENVELOPE_MISSING: The indexed entry no longer begins with a From separator.");
        return raw;
    }

    private static MboxVariant DetectVariant(Stream input, CancellationToken cancellationToken) {
        long visited = 0;
        long angles = 0;
        int matched = 0;
        bool prefix = true;
        int value;
        while ((value = input.ReadByte()) >= 0) {
            if ((visited++ & 0xFFFF) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (value == '\r' || value == '\n') { angles = 0; matched = 0; prefix = true; }
            else if (prefix && matched == 0 && value == '>') angles++;
            else if (prefix && angles > 1 && value == "From "[matched]) {
                if (++matched == 5) return MboxVariant.Mboxrd;
            } else prefix = false;
        }
        return MboxVariant.Mboxo;
    }

    private static void Unescape(Stream input, Stream output, MboxVariant variant, CancellationToken cancellationToken) {
        bool lineStart = true;
        long visited = 0;
        byte[] prefix = new byte[5];
        byte[] anglesBuffer = new byte[4096];
        for (int index = 0; index < anglesBuffer.Length; index++) anglesBuffer[index] = (byte)'>';
        int value;
        while ((value = input.ReadByte()) >= 0) {
            if ((visited++ & 0xFFFF) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (lineStart && value == '>') {
                long angles = 1;
                while ((value = input.ReadByte()) == '>') {
                    if ((visited++ & 0xFFFF) == 0) cancellationToken.ThrowIfCancellationRequested();
                    angles++;
                }
                int captured = 0;
                while (value >= 0 && captured < prefix.Length) {
                    prefix[captured++] = (byte)value;
                    if (value == '\r' || value == '\n' || captured == prefix.Length) break;
                    value = input.ReadByte();
                }
                bool from = captured == 5;
                for (int index = 0; from && index < 5; index++) from = prefix[index] == "From "[index];
                if (from && (angles == 1 || variant == MboxVariant.Mboxrd)) angles--;
                while (angles > 0) {
                    cancellationToken.ThrowIfCancellationRequested();
                    int count = (int)Math.Min(angles, anglesBuffer.Length);
                    output.Write(anglesBuffer, 0, count);
                    angles -= count;
                }
                output.Write(prefix, 0, captured);
                lineStart = captured > 0 && (prefix[captured - 1] == '\r' || prefix[captured - 1] == '\n');
            } else {
                output.WriteByte((byte)value);
                lineStart = value == '\r' || value == '\n';
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
    }
}
