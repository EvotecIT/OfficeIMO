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
        if (result.HasErrors && bytes.Length == 0) return result;
        OfficeFileCommit.WriteAllBytes(filePath, bytes);
        return result;
    }

    /// <summary>Writes a mailbox to a stream without closing it.</summary>
    public EmailWriteResult Write(EmailMailbox mailbox, Stream stream) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The stream must be writable.", nameof(stream));
        byte[] bytes = ToBytes(mailbox, out EmailWriteResult result, CancellationToken.None);
        if (result.HasErrors && bytes.Length == 0) return result;
        OfficeStreamWriter.WriteAllBytes(stream, bytes);
        return result;
    }

    /// <summary>Writes a mailbox to memory.</summary>
    public byte[] ToBytes(EmailMailbox mailbox) {
        byte[] bytes = ToBytes(mailbox, out EmailWriteResult result, CancellationToken.None);
        if (result.HasErrors) {
            EmailDiagnostic error = result.Diagnostics.First(diagnostic =>
                diagnostic.Severity == EmailDiagnosticSeverity.Error);
            throw new InvalidDataException(string.Concat("The mailbox could not be serialized: ", error.Code,
                ": ", error.Message));
        }
        return bytes;
    }

    /// <summary>
    /// Writes an entry sequence directly to a caller-owned stream while retaining at most one serialized message.
    /// A limit or enumeration failure can leave already completed entries in the destination.
    /// </summary>
    public EmailWriteResult WriteEntries(IEnumerable<EmailMailboxEntry> entries, Stream stream,
        CancellationToken cancellationToken = default) {
        if (entries == null) throw new ArgumentNullException(nameof(entries));
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The stream must be writable.", nameof(stream));
        var diagnostics = new List<EmailDiagnostic>();
        long bytesWritten = 0;
        int index = 0;
        foreach (EmailMailboxEntry entry in entries) {
            cancellationToken.ThrowIfCancellationRequested();
            byte[] bytes = SerializeEntry(entry, index, diagnostics, cancellationToken);
            if (bytes.Length == 0 && diagnostics.Any(diagnostic =>
                diagnostic.Severity == EmailDiagnosticSeverity.Error)) {
                return new EmailWriteResult(bytesWritten, diagnostics.AsReadOnly(), false);
            }
            long next = checked(bytesWritten + bytes.LongLength);
            stream.Write(bytes, 0, bytes.Length);
            bytesWritten = next;
            index++;
        }
        return new EmailWriteResult(bytesWritten, diagnostics.AsReadOnly(), false);
    }

    /// <summary>Asynchronously writes a mailbox file.</summary>
    public async Task<EmailWriteResult> WriteAsync(EmailMailbox mailbox, string filePath,
        CancellationToken cancellationToken = default) {
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        cancellationToken.ThrowIfCancellationRequested();
        byte[] bytes = ToBytes(mailbox, out EmailWriteResult result, cancellationToken);
        if (result.HasErrors && bytes.Length == 0) return result;
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
        if (result.HasErrors && bytes.Length == 0) return result;
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
                byte[] entry = SerializeEntry(mailbox.Messages[index], index, diagnostics, cancellationToken,
                    messageWriter);
                if (entry.Length == 0 && diagnostics.Any(diagnostic =>
                    diagnostic.Severity == EmailDiagnosticSeverity.Error)) {
                    result = new EmailWriteResult(0, diagnostics.AsReadOnly(), false);
                    return Array.Empty<byte>();
                }
                output.Write(entry, 0, entry.Length);
            }
            byte[] bytes = output.ToArray();
            cancellationToken.ThrowIfCancellationRequested();
            result = new EmailWriteResult(bytes.LongLength, diagnostics.AsReadOnly(), false);
            return bytes;
        }
    }

    private byte[] SerializeEntry(EmailMailboxEntry entry, int index, IList<EmailDiagnostic> diagnostics,
        CancellationToken cancellationToken, EmailDocumentWriter? existingWriter = null) {
        var messageWriter = existingWriter ?? new EmailDocumentWriter(_options.MessageOptions);
        byte[] eml;
        EmailWriteResult messageResult;
        using (var messageOutput = new MemoryStream()) {
            messageResult = messageWriter.Write(entry.Document, messageOutput, EmailFileFormat.Eml);
            eml = messageOutput.ToArray();
        }
        foreach (EmailDiagnostic diagnostic in messageResult.Diagnostics) {
            diagnostics.Add(new EmailDiagnostic(diagnostic.Code, diagnostic.Message, diagnostic.Severity,
                string.Concat("message[", index.ToString(CultureInfo.InvariantCulture), "]",
                    diagnostic.Location == null ? string.Empty : string.Concat("/", diagnostic.Location))));
        }
        if (messageResult.HasErrors && eml.Length == 0) return Array.Empty<byte>();

        byte[] fromBytes = CreateEnvelopeLine(entry, index, diagnostics);
        byte[] escaped = Escape(NormalizeLineEndings(eml), _options.Variant);
        using (var output = new MemoryStream(checked(fromBytes.Length + escaped.Length + 1))) {
            output.Write(fromBytes, 0, fromBytes.Length);
            output.Write(escaped, 0, escaped.Length);
            if (escaped.Length == 0 || escaped[escaped.Length - 1] != '\n') output.WriteByte((byte)'\n');
            cancellationToken.ThrowIfCancellationRequested();
            return output.ToArray();
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

    private static byte[] CreateEnvelopeLine(EmailMailboxEntry entry, int index,
        ICollection<EmailDiagnostic> diagnostics) {
        if (entry.RawFromLine != null) {
            if (IsSafeEnvelopeLine(entry.RawFromLine)) {
                return Encoding.ASCII.GetBytes(string.Concat(entry.RawFromLine, "\n"));
            }
            diagnostics.Add(new EmailDiagnostic("EMAIL_MBOX_RAW_ENVELOPE_INVALID",
                "The retained mbox separator was unsafe or not ASCII and was regenerated from structured metadata.",
                EmailDiagnosticSeverity.Warning,
                string.Concat("message[", index.ToString(CultureInfo.InvariantCulture), "]/envelope")));
        }

        string sender = entry.EnvelopeSender ?? entry.Document.From?.Address ?? "MAILER-DAEMON";
        DateTimeOffset date = entry.EnvelopeDate ?? entry.Document.Date ??
            new DateTimeOffset(1970, 1, 1, 0, 0, 0, TimeSpan.Zero);
        return Encoding.ASCII.GetBytes(string.Concat("From ", SanitizeSender(sender), " ",
            date.UtcDateTime.ToString("ddd MMM dd HH:mm:ss yyyy", CultureInfo.InvariantCulture), "\n"));
    }

    private static bool IsSafeEnvelopeLine(string value) {
        if (!value.StartsWith("From ", StringComparison.Ordinal)) return false;
        foreach (char character in value) {
            if (character >= 0x7F || (character < 0x20 && character != '\t')) return false;
        }
        return true;
    }

    private static string SanitizeSender(string sender) {
        return sender.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace(" ", string.Empty);
    }
}
