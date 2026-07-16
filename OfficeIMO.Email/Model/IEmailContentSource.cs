namespace OfficeIMO.Email;

/// <summary>
/// Supplies a fresh readable stream for content that should not be retained as a resident byte array.
/// Future mailbox stores can implement this contract without leaking store-specific handles into the email model.
/// </summary>
public interface IEmailContentSource {
    /// <summary>Known decoded content length, or null when the source cannot determine it cheaply.</summary>
    long? Length { get; }

    /// <summary>Opens a new readable stream. The caller owns and disposes the returned stream.</summary>
    Stream OpenRead();

    /// <summary>Asynchronously opens a new readable stream. The caller owns and disposes the returned stream.</summary>
    Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default);
}
