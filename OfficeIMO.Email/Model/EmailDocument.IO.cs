using OfficeIMO.Drawing.Internal;
namespace OfficeIMO.Email;

public sealed partial class EmailDocument {
    /// <summary>
    /// Loads one EML, MSG, or TNEF artifact with the default bounded policy.
    /// Use <see cref="EmailDocumentReader"/> when the caller also needs structured diagnostics.
    /// </summary>
    public static EmailDocument Load(string filePath, EmailReaderOptions? options = null,
        CancellationToken cancellationToken = default) =>
        GetDocumentOrThrow(new EmailDocumentReader(options ?? EmailReaderOptions.Default)
            .Read(filePath, cancellationToken));

    /// <summary>
    /// Loads one EML, MSG, or TNEF artifact from memory.
    /// Use <see cref="EmailDocumentReader"/> when the caller also needs structured diagnostics.
    /// </summary>
    public static EmailDocument Load(byte[] data, EmailReaderOptions? options = null,
        CancellationToken cancellationToken = default) =>
        GetDocumentOrThrow(new EmailDocumentReader(options ?? EmailReaderOptions.Default)
            .Read(data, cancellationToken));

    /// <summary>
    /// Loads one EML, MSG, or TNEF artifact without closing the stream. Seekable streams are read from the beginning
    /// and restored to their original position; non-seekable streams are read forward from their current position.
    /// Use <see cref="EmailDocumentReader"/> when the caller also needs structured diagnostics.
    /// </summary>
    public static EmailDocument Load(Stream stream, EmailReaderOptions? options = null,
        CancellationToken cancellationToken = default) =>
        GetDocumentOrThrow(new EmailDocumentReader(options ?? EmailReaderOptions.Default)
            .Read(stream, cancellationToken));

    /// <summary>Asynchronously loads one EML, MSG, or TNEF artifact with the default bounded policy.</summary>
    public static async Task<EmailDocument> LoadAsync(string filePath, EmailReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        EmailReadResult result = await new EmailDocumentReader(options ?? EmailReaderOptions.Default)
            .ReadAsync(filePath, cancellationToken).ConfigureAwait(false);
        return GetDocumentOrThrow(result);
    }

    /// <summary>
    /// Asynchronously loads an artifact without closing the stream. Seekable streams are read from the beginning
    /// and restored to their original position; non-seekable streams are read forward.
    /// </summary>
    public static async Task<EmailDocument> LoadAsync(Stream stream, EmailReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        EmailReadResult result = await new EmailDocumentReader(options ?? EmailReaderOptions.Default)
            .ReadAsync(stream, cancellationToken).ConfigureAwait(false);
        return GetDocumentOrThrow(result);
    }

    /// <summary>Saves the document as EML, MSG, or TNEF, inferred from the destination filename.</summary>
    public EmailWriteResult Save(string filePath, EmailWriterOptions? options = null) =>
        Save(filePath, InferOutputFormat(filePath), options);

    /// <summary>Saves the document in the explicitly selected artifact format.</summary>
    public EmailWriteResult Save(string filePath, EmailFileFormat format, EmailWriterOptions? options = null) {
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        EmailDocumentWriter writer = new EmailDocumentWriter(options ?? EmailWriterOptions.Default);
        byte[] data = writer.ToBytes(this, format, out EmailWriteResult result);
        EnsureWriteSucceeded(result);
        OfficeFileCommit.WriteAllBytes(filePath, data);
        return result;
    }

    /// <summary>Saves the document to a stream without closing it.</summary>
    public EmailWriteResult Save(Stream stream, EmailFileFormat format = EmailFileFormat.Eml,
        EmailWriterOptions? options = null) {
        EmailDocumentWriter writer = new EmailDocumentWriter(options ?? EmailWriterOptions.Default);
        byte[] data = writer.ToBytes(this, format, out EmailWriteResult result);
        EnsureWriteSucceeded(result);
        OfficeStreamWriter.WriteAllBytes(stream, data);
        return result;
    }

    /// <summary>Asynchronously saves the document, inferring the format from the destination filename.</summary>
    public Task<EmailWriteResult> SaveAsync(string filePath, EmailWriterOptions? options = null,
        CancellationToken cancellationToken = default) =>
        SaveAsync(filePath, InferOutputFormat(filePath), options, cancellationToken);

    /// <summary>Asynchronously saves the document in the explicitly selected artifact format.</summary>
    public async Task<EmailWriteResult> SaveAsync(string filePath, EmailFileFormat format,
        EmailWriterOptions? options = null, CancellationToken cancellationToken = default) {
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        cancellationToken.ThrowIfCancellationRequested();
        EmailDocumentWriter writer = new EmailDocumentWriter(options ?? EmailWriterOptions.Default);
        byte[] data = writer.ToBytes(this, format, out EmailWriteResult result);
        EnsureWriteSucceeded(result);
        await OfficeFileCommit.WriteAllBytesAsync(filePath, data, cancellationToken: cancellationToken)
            .ConfigureAwait(false);
        return result;
    }

    /// <summary>Asynchronously saves the document to a stream without closing it.</summary>
    public async Task<EmailWriteResult> SaveAsync(Stream stream, EmailFileFormat format = EmailFileFormat.Eml,
        EmailWriterOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        EmailDocumentWriter writer = new EmailDocumentWriter(options ?? EmailWriterOptions.Default);
        byte[] data = writer.ToBytes(this, format, out EmailWriteResult result);
        EnsureWriteSucceeded(result);
        cancellationToken.ThrowIfCancellationRequested();
        await OfficeStreamWriter.WriteAllBytesAsync(stream, data, cancellationToken).ConfigureAwait(false);
        return result;
    }

    /// <summary>Serializes the document to memory.</summary>
    public byte[] ToBytes(EmailFileFormat format = EmailFileFormat.Eml, EmailWriterOptions? options = null) {
        EmailDocumentWriter writer = new EmailDocumentWriter(options ?? EmailWriterOptions.Default);
        byte[] data = writer.ToBytes(this, format, out EmailWriteResult result);
        EnsureWriteSucceeded(result);
        return data;
    }

    /// <summary>Serializes the document to a new writable memory stream positioned at the beginning.</summary>
    public MemoryStream ToStream(EmailFileFormat format = EmailFileFormat.Eml, EmailWriterOptions? options = null) =>
        new MemoryStream(ToBytes(format, options));

    private static EmailDocument GetDocumentOrThrow(EmailReadResult result) {
        if (!result.HasErrors) return result.Document;
        throw CreateDiagnosticException("The email artifact could not be loaded", result.Diagnostics);
    }

    private static void EnsureWriteSucceeded(EmailWriteResult result) {
        if (result.HasErrors) {
            throw CreateDiagnosticException("The email artifact could not be saved", result.Diagnostics);
        }
    }

    private static InvalidDataException CreateDiagnosticException(string message,
        IReadOnlyList<EmailDiagnostic> diagnostics) {
        foreach (EmailDiagnostic diagnostic in diagnostics) {
            if (diagnostic.Severity == EmailDiagnosticSeverity.Error) {
                return new InvalidDataException(string.Concat(message, ": ", diagnostic.Code, ": ", diagnostic.Message));
            }
        }
        return new InvalidDataException(message + ".");
    }

    private static EmailFileFormat InferOutputFormat(string filePath) {
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        string fileName = Path.GetFileName(filePath);
        string extension = Path.GetExtension(fileName);
        if (extension.Equals(".eml", StringComparison.OrdinalIgnoreCase) ||
            extension.Equals(".mime", StringComparison.OrdinalIgnoreCase)) {
            return EmailFileFormat.Eml;
        }
        if (extension.Equals(".msg", StringComparison.OrdinalIgnoreCase)) return EmailFileFormat.OutlookMsg;
        if (extension.Equals(".tnef", StringComparison.OrdinalIgnoreCase) ||
            fileName.Equals("winmail.dat", StringComparison.OrdinalIgnoreCase)) {
            return EmailFileFormat.Tnef;
        }
        throw new NotSupportedException(
            "Cannot infer the email format from the destination filename. Use .eml, .msg, .tnef, or winmail.dat, or call Save with an explicit EmailFileFormat.");
    }
}
