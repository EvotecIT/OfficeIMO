namespace OfficeIMO.Email;

/// <summary>Overrides attachment content sources for one serialization operation.</summary>
internal sealed class EmailAttachmentStreamScope : IDisposable {
    private static readonly AsyncLocal<EmailAttachmentStreamScope?> CurrentScope =
        new AsyncLocal<EmailAttachmentStreamScope?>();
    private readonly EmailAttachmentStreamScope? _previous;
    private readonly IReadOnlyDictionary<EmailAttachment, IEmailContentSource> _sources;
    private bool _disposed;

    private EmailAttachmentStreamScope(IReadOnlyDictionary<EmailAttachment, IEmailContentSource> sources) {
        _sources = sources;
        _previous = CurrentScope.Value;
        CurrentScope.Value = this;
    }

    internal static IDisposable Begin(IReadOnlyDictionary<EmailAttachment, IEmailContentSource> sources) {
        if (sources == null) throw new ArgumentNullException(nameof(sources));
        return new EmailAttachmentStreamScope(sources);
    }

    internal static bool HasStagedContent(EmailAttachment attachment) =>
        CurrentScope.Value?._sources.ContainsKey(attachment) == true;

    internal static Stream OpenRead(EmailAttachment attachment) {
        if (attachment == null) throw new ArgumentNullException(nameof(attachment));
        if (CurrentScope.Value?._sources.TryGetValue(attachment, out IEmailContentSource? source) == true) {
            Stream stream = source.OpenRead();
            if (stream == null || !stream.CanRead) {
                stream?.Dispose();
                throw new InvalidDataException("The staged attachment source did not return a readable stream.");
            }
            return stream;
        }
        return attachment.OpenContentStream();
    }

    internal static long? GetLength(EmailAttachment attachment) {
        if (attachment.Content != null) return attachment.Content.LongLength;
        if (CurrentScope.Value?._sources.TryGetValue(attachment, out IEmailContentSource? source) == true) {
            return source.Length;
        }
        return attachment.ContentSource?.Length ?? (attachment.Length > 0 ? attachment.Length : (long?)null);
    }

    public void Dispose() {
        if (_disposed) return;
        if (!ReferenceEquals(CurrentScope.Value, this)) {
            throw new InvalidOperationException("Attachment stream scopes must be disposed in reverse order.");
        }
        CurrentScope.Value = _previous;
        _disposed = true;
    }
}
