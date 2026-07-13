using OfficeIMO.Drawing.Internal;
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
        byte[] data = ToBytes(document, format, out EmailWriteResult result);
        OfficeFileCommit.WriteAllBytes(filePath, data);
        return result;
    }

    /// <summary>Writes an artifact to a stream without closing it.</summary>
    public EmailWriteResult Write(EmailDocument document, Stream stream, EmailFileFormat format = EmailFileFormat.Eml) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The stream must be writable.", nameof(stream));
        byte[] data = ToBytes(document, format, out EmailWriteResult result);
        OfficeStreamWriter.WriteAllBytes(stream, data);
        return result;
    }

    /// <summary>Writes an artifact to memory.</summary>
    public byte[] ToBytes(EmailDocument document, EmailFileFormat format = EmailFileFormat.Eml) {
        return ToBytes(document, format, out _);
    }

    /// <summary>Asynchronously writes an artifact to a file.</summary>
    public async Task<EmailWriteResult> WriteAsync(EmailDocument document, string filePath,
        EmailFileFormat format = EmailFileFormat.Eml, CancellationToken cancellationToken = default) {
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        cancellationToken.ThrowIfCancellationRequested();
        byte[] data = ToBytes(document, format, out EmailWriteResult result);
        cancellationToken.ThrowIfCancellationRequested();
        await OfficeFileCommit.WriteAllBytesAsync(filePath, data, cancellationToken: cancellationToken)
            .ConfigureAwait(false);
        return result;
    }

    /// <summary>Asynchronously writes an artifact to a stream without closing it.</summary>
    public async Task<EmailWriteResult> WriteAsync(EmailDocument document, Stream stream,
        EmailFileFormat format = EmailFileFormat.Eml, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The stream must be writable.", nameof(stream));
        cancellationToken.ThrowIfCancellationRequested();
        byte[] data = ToBytes(document, format, out EmailWriteResult result);
        cancellationToken.ThrowIfCancellationRequested();
        await OfficeStreamWriter.WriteAllBytesAsync(stream, data, cancellationToken).ConfigureAwait(false);
        return result;
    }

    internal byte[] ToBytes(EmailDocument document, EmailFileFormat format, out EmailWriteResult result) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (format != EmailFileFormat.Eml && format != EmailFileFormat.OutlookMsg && format != EmailFileFormat.Tnef) {
            throw new NotSupportedException("The requested email artifact format cannot be serialized.");
        }

        List<EmailDiagnostic> diagnostics = new List<EmailDiagnostic>();
        if (_options.UsePreservedRawSource && document.Format == format && document.RawSource != null) {
            byte[]? baseline = document.RawSourceModelFingerprint;
            if (baseline != null && EmailDocumentStateFingerprint.Matches(document, baseline)) {
                EnsureOutputLimit(document.RawSource.LongLength);
                byte[] preserved = (byte[])document.RawSource.Clone();
                result = new EmailWriteResult(preserved.LongLength, diagnostics.AsReadOnly(), true);
                return preserved;
            }
            diagnostics.Add(new EmailDiagnostic("EMAIL_RAW_SOURCE_SKIPPED_MODEL_CHANGED",
                "The preserved source was not reused because the email model changed after reading or could not be verified as unchanged.",
                EmailDiagnosticSeverity.Warning));
        }

        EmailOutputPreflight.EnsurePayloadsFit(document, _options.MaxOutputBytes);
        byte[] data = format == EmailFileFormat.Eml
            ? MimeWriter.Write(document, _options, diagnostics)
            : format == EmailFileFormat.OutlookMsg
                ? MsgWriter.Write(document, _options, diagnostics)
                : TnefWriter.Write(document, _options, diagnostics);
        EnsureOutputLimit(data.LongLength);
        result = new EmailWriteResult(data.LongLength, diagnostics.AsReadOnly(), false);
        return data;
    }

    private void EnsureOutputLimit(long length) {
        if (length > _options.MaxOutputBytes) {
            throw new EmailLimitExceededException(nameof(EmailWriterOptions.MaxOutputBytes), length, _options.MaxOutputBytes);
        }
    }
}
