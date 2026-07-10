namespace OfficeIMO.Email;

/// <summary>Serializes <see cref="EmailDocument"/> instances into deterministic email artifacts.</summary>
public sealed class EmailDocumentWriter {
    private readonly EmailWriterOptions _options;

    /// <summary>Creates a writer with the default deterministic policy.</summary>
    public EmailDocumentWriter() : this(EmailWriterOptions.Default) { }

    /// <summary>Creates a writer with an immutable serialization policy.</summary>
    public EmailDocumentWriter(EmailWriterOptions options) {
        _options = options ?? throw new ArgumentNullException(nameof(options));
    }

    /// <summary>Writer policy used by this instance.</summary>
    public EmailWriterOptions Options => _options;

    /// <summary>Writes an artifact to a file.</summary>
    public EmailWriteResult Write(EmailDocument document, string filePath, EmailFileFormat format = EmailFileFormat.Eml) {
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        byte[] data = WriteToBytes(document, format, out EmailWriteResult result);
        File.WriteAllBytes(filePath, data);
        return result;
    }

    /// <summary>Writes an artifact to a stream without closing it.</summary>
    public EmailWriteResult Write(EmailDocument document, Stream stream, EmailFileFormat format = EmailFileFormat.Eml) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The stream must be writable.", nameof(stream));
        byte[] data = WriteToBytes(document, format, out EmailWriteResult result);
        stream.Write(data, 0, data.Length);
        return result;
    }

    /// <summary>Writes an artifact to memory.</summary>
    public byte[] WriteToBytes(EmailDocument document, EmailFileFormat format = EmailFileFormat.Eml) {
        return WriteToBytes(document, format, out _);
    }

    /// <summary>Asynchronously writes an artifact to a stream without closing it.</summary>
    public async Task<EmailWriteResult> WriteAsync(EmailDocument document, Stream stream,
        EmailFileFormat format = EmailFileFormat.Eml, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The stream must be writable.", nameof(stream));
        byte[] data = WriteToBytes(document, format, out EmailWriteResult result);
        await stream.WriteAsync(data, 0, data.Length, cancellationToken).ConfigureAwait(false);
        return result;
    }

    private byte[] WriteToBytes(EmailDocument document, EmailFileFormat format, out EmailWriteResult result) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (format != EmailFileFormat.Eml) throw new NotSupportedException("Only EML serialization is available in this delivery slice.");

        if (_options.UsePreservedRawSource && document.Format == format && document.RawSource != null) {
            byte[] preserved = (byte[])document.RawSource.Clone();
            result = new EmailWriteResult(preserved.LongLength, Array.Empty<EmailDiagnostic>(), true);
            return preserved;
        }

        List<EmailDiagnostic> diagnostics = new List<EmailDiagnostic>();
        byte[] data = MimeWriter.Write(document, _options, diagnostics);
        result = new EmailWriteResult(data.LongLength, diagnostics.AsReadOnly(), false);
        return data;
    }
}
