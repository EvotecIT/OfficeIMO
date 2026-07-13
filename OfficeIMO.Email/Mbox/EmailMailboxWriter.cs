using OfficeIMO.Drawing.Internal;
namespace OfficeIMO.Email;

/// <summary>Writes deterministic mboxo or mboxrd mailbox aggregates.</summary>
public sealed class EmailMailboxWriter {
    private readonly EmailMailboxWriterOptions _options;

    /// <summary>Creates a writer with the default mboxrd policy.</summary>
    public EmailMailboxWriter() : this(EmailMailboxWriterOptions.Default) { }

    /// <summary>Creates a writer with an immutable policy.</summary>
    public EmailMailboxWriter(EmailMailboxWriterOptions options) {
        _options = options ?? throw new ArgumentNullException(nameof(options));
    }

    /// <summary>Writer policy used by this instance.</summary>
    public EmailMailboxWriterOptions Options => _options;

    /// <summary>Writes a mailbox to a file.</summary>
    public EmailWriteResult Write(EmailMailbox mailbox, string filePath) {
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        byte[] bytes = ToBytes(mailbox, out EmailWriteResult result, CancellationToken.None);
        OfficeFileCommit.WriteAllBytes(filePath, bytes);
        return result;
    }

    /// <summary>Writes a mailbox to a stream without closing it.</summary>
    public EmailWriteResult Write(EmailMailbox mailbox, Stream stream) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The stream must be writable.", nameof(stream));
        byte[] bytes = ToBytes(mailbox, out EmailWriteResult result, CancellationToken.None);
        OfficeStreamWriter.WriteAllBytes(stream, bytes);
        return result;
    }

    /// <summary>Writes a mailbox to memory.</summary>
    public byte[] ToBytes(EmailMailbox mailbox) => ToBytes(mailbox, out _, CancellationToken.None);

    /// <summary>Asynchronously writes a mailbox file.</summary>
    public async Task<EmailWriteResult> WriteAsync(EmailMailbox mailbox, string filePath,
        CancellationToken cancellationToken = default) {
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        cancellationToken.ThrowIfCancellationRequested();
        byte[] bytes = ToBytes(mailbox, out EmailWriteResult result, cancellationToken);
        await OfficeFileCommit.WriteAllBytesAsync(filePath, bytes, cancellationToken: cancellationToken)
            .ConfigureAwait(false);
        return result;
    }

    /// <summary>Asynchronously writes a mailbox without closing the stream.</summary>
    public async Task<EmailWriteResult> WriteAsync(EmailMailbox mailbox, Stream stream,
        CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The stream must be writable.", nameof(stream));
        cancellationToken.ThrowIfCancellationRequested();
        byte[] bytes = ToBytes(mailbox, out EmailWriteResult result, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        await OfficeStreamWriter.WriteAllBytesAsync(stream, bytes, cancellationToken).ConfigureAwait(false);
        return result;
    }

    private byte[] ToBytes(EmailMailbox mailbox, out EmailWriteResult result, CancellationToken cancellationToken) {
        if (mailbox == null) throw new ArgumentNullException(nameof(mailbox));
        cancellationToken.ThrowIfCancellationRequested();
        var diagnostics = new List<EmailDiagnostic>();
        var messageWriter = new EmailDocumentWriter(_options.MessageOptions);
        using (EmailBoundedMemoryStream output = new EmailBoundedMemoryStream(
                   _options.MessageOptions.MaxOutputBytes)) {
            for (int index = 0; index < mailbox.Messages.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                EmailMailboxEntry entry = mailbox.Messages[index];
                string sender = entry.EnvelopeSender ?? entry.Document.From?.Address ?? "MAILER-DAEMON";
                DateTimeOffset date = entry.EnvelopeDate ?? entry.Document.Date ??
                    new DateTimeOffset(1970, 1, 1, 0, 0, 0, TimeSpan.Zero);
                string fromLine = string.Concat("From ", SanitizeSender(sender), " ",
                    date.UtcDateTime.ToString("ddd MMM dd HH:mm:ss yyyy", CultureInfo.InvariantCulture), "\n");
                byte[] fromBytes = Encoding.ASCII.GetBytes(fromLine);
                output.Write(fromBytes, 0, fromBytes.Length);
                byte[] eml;
                using (MemoryStream messageOutput = new MemoryStream()) {
                    EmailWriteResult messageResult = messageWriter.Write(entry.Document, messageOutput, EmailFileFormat.Eml);
                    diagnostics.AddRange(messageResult.Diagnostics);
                    eml = messageOutput.ToArray();
                }
                byte[] normalized = NormalizeLineEndings(eml);
                byte[] escaped = Escape(normalized, _options.Variant);
                output.Write(escaped, 0, escaped.Length);
                if (escaped.Length == 0 || escaped[escaped.Length - 1] != '\n') output.WriteByte((byte)'\n');
            }
            byte[] bytes = output.ToArray();
            cancellationToken.ThrowIfCancellationRequested();
            result = new EmailWriteResult(bytes.LongLength, diagnostics.AsReadOnly(), false);
            return bytes;
        }
    }

    private static byte[] Escape(byte[] message, MboxVariant variant) {
        using (MemoryStream output = new MemoryStream(message.Length)) {
            int lineStart = 0;
            while (lineStart < message.Length) {
                int lineEnd = lineStart;
                while (lineEnd < message.Length && message[lineEnd] != '\n') lineEnd++;
                int position = lineStart;
                if (variant == MboxVariant.Mboxrd) while (position < lineEnd && message[position] == '>') position++;
                if (StartsWith(message, position, "From ")) output.WriteByte((byte)'>');
                output.Write(message, lineStart, lineEnd - lineStart);
                if (lineEnd < message.Length) output.WriteByte((byte)'\n');
                lineStart = lineEnd + 1;
            }
            return output.ToArray();
        }
    }

    private static byte[] NormalizeLineEndings(byte[] input) {
        using (MemoryStream output = new MemoryStream(input.Length)) {
            for (int index = 0; index < input.Length; index++) {
                if (input[index] == '\r') {
                    if (index + 1 < input.Length && input[index + 1] == '\n') index++;
                    output.WriteByte((byte)'\n');
                } else {
                    output.WriteByte(input[index]);
                }
            }
            return output.ToArray();
        }
    }

    private static bool StartsWith(byte[] data, int offset, string value) {
        if (offset < 0 || offset + value.Length > data.Length) return false;
        for (int index = 0; index < value.Length; index++) if (data[offset + index] != value[index]) return false;
        return true;
    }

    private static string SanitizeSender(string sender) {
        return sender.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace(" ", string.Empty);
    }
}
