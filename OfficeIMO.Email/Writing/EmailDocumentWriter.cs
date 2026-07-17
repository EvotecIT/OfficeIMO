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
        WritePreparation preparation = Prepare(document, format);
        if (!preparation.CanWrite) return preparation.CreateBlockedResult();
        EmailWriteResult? result = null;
        OfficeFileCommit.Write(filePath, stream => result = WritePrepared(preparation, stream));
        return result!;
    }

    /// <summary>Writes an artifact to a stream without closing it.</summary>
    public EmailWriteResult Write(EmailDocument document, Stream stream, EmailFileFormat format = EmailFileFormat.Eml) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The stream must be writable.", nameof(stream));
        WritePreparation preparation = Prepare(document, format);
        if (!preparation.CanWrite) return preparation.CreateBlockedResult();
        EmailWriteResult? result = null;
        OfficeStreamWriter.Write(stream, output => result = WritePrepared(preparation, output));
        return result!;
    }

    /// <summary>Writes an artifact to memory.</summary>
    public byte[] ToBytes(EmailDocument document, EmailFileFormat format = EmailFileFormat.Eml) {
        byte[] data = ToBytes(document, format, out EmailWriteResult result);
        ThrowIfBlocked(result);
        return data;
    }

    /// <summary>Asynchronously writes an artifact to a file.</summary>
    public async Task<EmailWriteResult> WriteAsync(EmailDocument document, string filePath,
        EmailFileFormat format = EmailFileFormat.Eml, CancellationToken cancellationToken = default) {
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        cancellationToken.ThrowIfCancellationRequested();
        WritePreparation preparation = Prepare(document, format);
        if (!preparation.CanWrite) return preparation.CreateBlockedResult();
        using (EmailAttachmentStaging staging = preparation.PreservedSource != null
                   ? EmailAttachmentStaging.CreateEmpty()
                   : await EmailAttachmentStaging.CreateAsync(
                       document, _options.MaxOutputBytes, cancellationToken).ConfigureAwait(false))
        using (staging.EnterScope()) {
            EmailWriteResult? result = null;
            await OfficeFileCommit.WriteAsync(filePath,
                async (output, token) => result = await WritePreparedAsync(preparation, output, token)
                    .ConfigureAwait(false),
                cancellationToken: cancellationToken).ConfigureAwait(false);
            return result!;
        }
    }

    /// <summary>Asynchronously writes an artifact to a stream without closing it.</summary>
    public async Task<EmailWriteResult> WriteAsync(EmailDocument document, Stream stream,
        EmailFileFormat format = EmailFileFormat.Eml, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The stream must be writable.", nameof(stream));
        cancellationToken.ThrowIfCancellationRequested();
        WritePreparation preparation = Prepare(document, format);
        if (!preparation.CanWrite) return preparation.CreateBlockedResult();
        using (EmailAttachmentStaging staging = preparation.PreservedSource != null
                   ? EmailAttachmentStaging.CreateEmpty()
                   : await EmailAttachmentStaging.CreateAsync(
                       document, _options.MaxOutputBytes, cancellationToken).ConfigureAwait(false))
        using (staging.EnterScope()) {
            EmailWriteResult? result = null;
            await OfficeStreamWriter.WriteAsync(stream,
                async (output, token) => result = await WritePreparedAsync(preparation, output, token)
                    .ConfigureAwait(false), cancellationToken).ConfigureAwait(false);
            return result!;
        }
    }

    internal byte[] ToBytes(EmailDocument document, EmailFileFormat format, out EmailWriteResult result) {
        WritePreparation preparation = Prepare(document, format);
        if (!preparation.CanWrite) {
            result = preparation.CreateBlockedResult();
            return Array.Empty<byte>();
        }
        using (var output = new MemoryStream()) {
            result = WritePrepared(preparation, output);
            return output.ToArray();
        }
    }

    private WritePreparation Prepare(EmailDocument document, EmailFileFormat format) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (format != EmailFileFormat.Eml && format != EmailFileFormat.OutlookMsg &&
            format != EmailFileFormat.OutlookTemplate && format != EmailFileFormat.Tnef) {
            throw new NotSupportedException("The requested email artifact format cannot be serialized.");
        }

        List<EmailDiagnostic> diagnostics = new List<EmailDiagnostic>();
        bool protectedPassThrough = EmailConversionAnalyzer.CanPassThroughProtectedSource(document, format);
        if ((_options.UsePreservedRawSource || protectedPassThrough) && document.Format == format && document.RawSource != null) {
            byte[]? baseline = document.RawSourceModelFingerprint;
            if (baseline != null && EmailDocumentStateFingerprint.Matches(document, baseline)) {
                EnsureOutputLimit(document.RawSource.LongLength);
                return new WritePreparation(document, format, diagnostics, document.RawSource);
            }
            diagnostics.Add(new EmailDiagnostic("EMAIL_RAW_SOURCE_SKIPPED_MODEL_CHANGED",
                "The preserved source was not reused because the email model changed after reading or could not be verified as unchanged.",
                EmailDiagnosticSeverity.Warning));
        }

        EmailConversionReport conversion = EmailConversionAnalyzer.Analyze(document, format, _options);
        diagnostics.AddRange(conversion.Diagnostics);
        if (!conversion.CanWrite) {
            return new WritePreparation(document, format, diagnostics, preservedSource: null, canWrite: false);
        }

        EmailOutputPreflight.EnsurePayloadsFit(document, format, _options.MaxOutputBytes);
        return new WritePreparation(document, format, diagnostics, preservedSource: null);
    }

    private EmailWriteResult WritePrepared(WritePreparation preparation, Stream output) {
        using (var bounded = new EmailBoundedWriteStream(output, _options.MaxOutputBytes)) {
            if (preparation.PreservedSource != null) {
                bounded.Write(preparation.PreservedSource, 0, preparation.PreservedSource.Length);
            } else if (preparation.Format == EmailFileFormat.Eml) {
                MimeWriter.Write(bounded, preparation.Document, _options, preparation.Diagnostics);
            } else if (preparation.Format == EmailFileFormat.Tnef) {
                TnefWriter.Write(bounded, preparation.Document, _options, preparation.Diagnostics);
            } else {
                MsgWriter.Write(bounded, preparation.Document, _options, preparation.Diagnostics,
                    preparation.Format == EmailFileFormat.OutlookTemplate);
            }
            return new EmailWriteResult(bounded.BytesWritten, preparation.Diagnostics.AsReadOnly(),
                preparation.PreservedSource != null);
        }
    }

    private async Task<EmailWriteResult> WritePreparedAsync(WritePreparation preparation, Stream output,
        CancellationToken cancellationToken) {
        if (preparation.PreservedSource != null) {
            using (var bounded = new EmailBoundedWriteStream(output, _options.MaxOutputBytes)) {
                await bounded.WriteAsync(preparation.PreservedSource, 0, preparation.PreservedSource.Length,
                    cancellationToken).ConfigureAwait(false);
                return new EmailWriteResult(bounded.BytesWritten, preparation.Diagnostics.AsReadOnly(), true);
            }
        }

        long bytesWritten = await EmailAsyncWritePipeline.RunAsync(output, _options.MaxOutputBytes,
            producer => SerializeGenerated(preparation, producer), cancellationToken).ConfigureAwait(false);
        return new EmailWriteResult(bytesWritten, preparation.Diagnostics.AsReadOnly(), false);
    }

    private void SerializeGenerated(WritePreparation preparation, Stream output) {
        if (preparation.Format == EmailFileFormat.Eml) {
            MimeWriter.Write(output, preparation.Document, _options, preparation.Diagnostics);
            return;
        }
        if (preparation.Format == EmailFileFormat.Tnef) {
            TnefWriter.Write(output, preparation.Document, _options, preparation.Diagnostics);
            return;
        }
        if (preparation.Format == EmailFileFormat.OutlookMsg ||
            preparation.Format == EmailFileFormat.OutlookTemplate) {
            MsgWriter.Write(output, preparation.Document, _options, preparation.Diagnostics,
                preparation.Format == EmailFileFormat.OutlookTemplate);
            return;
        }
        throw new NotSupportedException();
    }

    private void EnsureOutputLimit(long length) {
        if (length > _options.MaxOutputBytes) {
            throw new EmailLimitExceededException(nameof(EmailWriterOptions.MaxOutputBytes), length, _options.MaxOutputBytes);
        }
    }

    /// <summary>Analyzes known fidelity implications without producing output.</summary>
    public EmailConversionReport AnalyzeConversion(EmailDocument document,
        EmailFileFormat format = EmailFileFormat.Eml) {
        if (format != EmailFileFormat.Eml && format != EmailFileFormat.OutlookMsg &&
            format != EmailFileFormat.OutlookTemplate && format != EmailFileFormat.Tnef) {
            throw new NotSupportedException("The requested email artifact format cannot be serialized.");
        }
        return EmailConversionAnalyzer.Analyze(document, format, _options);
    }

    private static void ThrowIfBlocked(EmailWriteResult result) {
        if (!result.HasErrors) return;
        EmailDiagnostic diagnostic = result.Diagnostics.First(item => item.Severity == EmailDiagnosticSeverity.Error);
        throw new InvalidDataException(string.Concat("The email artifact could not be serialized: ",
            diagnostic.Code, ": ", diagnostic.Message));
    }

    private sealed class WritePreparation {
        internal WritePreparation(EmailDocument document, EmailFileFormat format,
            List<EmailDiagnostic> diagnostics, byte[]? preservedSource, bool canWrite = true) {
            Document = document;
            Format = format;
            Diagnostics = diagnostics;
            PreservedSource = preservedSource;
            CanWrite = canWrite;
        }

        internal EmailDocument Document { get; }
        internal EmailFileFormat Format { get; }
        internal List<EmailDiagnostic> Diagnostics { get; }
        internal byte[]? PreservedSource { get; }
        internal bool CanWrite { get; }

        internal EmailWriteResult CreateBlockedResult() =>
            new EmailWriteResult(0, Diagnostics.AsReadOnly(), false);
    }
}
