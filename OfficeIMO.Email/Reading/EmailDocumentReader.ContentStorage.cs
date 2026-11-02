namespace OfficeIMO.Email;

/// <summary>Controls the lifetime and storage of retained attachment payloads.</summary>
public enum EmailReadContentStorage {
    /// <summary>Retains payloads in memory.</summary>
    Memory,
    /// <summary>Uses temporary files owned by the disposable read result.</summary>
    TemporaryFiles
}

public sealed partial class EmailDocumentReader {
    /// <summary>Reads a complete document using explicit payload storage. Dispose temporary-file results after use.</summary>
    public EmailReadResult ReadDocument(Stream stream, EmailReadContentStorage contentStorage, string? sourceName = null, CancellationToken cancellationToken = default) =>
        contentStorage switch {
            EmailReadContentStorage.Memory => ReadDocument(stream, sourceName, cancellationToken),
            EmailReadContentStorage.TemporaryFiles => ReadFileBacked(stream, sourceName, cancellationToken),
            _ => throw new ArgumentOutOfRangeException(nameof(contentStorage))
        };

    /// <summary>Reads a complete file using explicit payload storage.</summary>
    public EmailReadResult ReadDocument(string filePath, EmailReadContentStorage contentStorage, CancellationToken cancellationToken = default) =>
        contentStorage switch {
            EmailReadContentStorage.Memory => ReadDocument(filePath, cancellationToken),
            EmailReadContentStorage.TemporaryFiles => ReadFileBacked(filePath, cancellationToken),
            _ => throw new ArgumentOutOfRangeException(nameof(contentStorage))
        };

    /// <summary>Reads a complete document asynchronously using explicit payload storage.</summary>
    public Task<EmailReadResult> ReadDocumentAsync(Stream stream, EmailReadContentStorage contentStorage, string? sourceName = null, CancellationToken cancellationToken = default) =>
        contentStorage switch {
            EmailReadContentStorage.Memory => ReadDocumentAsync(stream, sourceName, cancellationToken),
            EmailReadContentStorage.TemporaryFiles => ReadFileBackedAsync(stream, sourceName, cancellationToken),
            _ => throw new ArgumentOutOfRangeException(nameof(contentStorage))
        };

    /// <summary>Reads a complete file asynchronously using explicit payload storage.</summary>
    public Task<EmailReadResult> ReadDocumentAsync(string filePath, EmailReadContentStorage contentStorage, CancellationToken cancellationToken = default) =>
        contentStorage switch {
            EmailReadContentStorage.Memory => ReadDocumentAsync(filePath, cancellationToken),
            EmailReadContentStorage.TemporaryFiles => ReadFileBackedAsync(filePath, cancellationToken),
            _ => throw new ArgumentOutOfRangeException(nameof(contentStorage))
        };
}
