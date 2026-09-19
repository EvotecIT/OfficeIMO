namespace OfficeIMO.Email;

internal sealed class MimeWriterState : IDisposable {
    private readonly HashSet<EmailDocument> _activeDocuments = new HashSet<EmailDocument>();
    private readonly Dictionary<EmailAttachment, Stream> _preparedAttachmentStreams =
        new Dictionary<EmailAttachment, Stream>(AttachmentReferenceComparer.Instance);
    private string? _temporaryDirectoryPath;
    private bool _disposed;

    internal MimeWriterState(EmailWriterOptions options, IList<EmailDiagnostic> diagnostics) {
        Options = options;
        Diagnostics = diagnostics;
    }

    internal EmailWriterOptions Options { get; }

    internal IList<EmailDiagnostic> Diagnostics { get; }

    internal void Enter(EmailDocument document, int depth) {
        if (depth > Options.MaxNestedMessageDepth) {
            throw new InvalidOperationException("The embedded-message write depth exceeds the configured maximum.");
        }
        if (!_activeDocuments.Add(document)) throw new InvalidOperationException("The embedded-message graph contains a cycle.");
    }

    internal void Exit(EmailDocument document) {
        _activeDocuments.Remove(document);
    }

    internal Stream PrepareAttachmentStream(EmailAttachment attachment) {
        if (_preparedAttachmentStreams.TryGetValue(attachment, out Stream? prepared)) {
            prepared.Position = 0;
            return prepared;
        }

        Stream source = EmailAttachmentStreamScope.OpenRead(attachment);
        if (source.CanSeek) {
            source.Position = 0;
            _preparedAttachmentStreams.Add(attachment, source);
            return source;
        }

        string directory = EnsureTemporaryDirectory();
        string path = Path.Combine(directory,
            string.Concat(Guid.NewGuid().ToString("N"), ".content"));
        try {
            long copied = 0;
            using (source)
            using (var destination = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read,
                       81920, FileOptions.SequentialScan)) {
                var buffer = new byte[81920];
                while (true) {
                    int read = source.Read(buffer, 0, buffer.Length);
                    if (read == 0) break;
                    copied = checked(copied + read);
                    if (copied > Options.MaxOutputBytes) {
                        throw new EmailLimitExceededException(nameof(EmailWriterOptions.MaxOutputBytes),
                            copied, Options.MaxOutputBytes);
                    }
                    destination.Write(buffer, 0, read);
                }
            }
            prepared = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read,
                81920, FileOptions.SequentialScan);
            _preparedAttachmentStreams.Add(attachment, prepared);
            return prepared;
        } catch {
            try { if (File.Exists(path)) File.Delete(path); } catch { }
            throw;
        }
    }

    internal Stream OpenAttachmentStream(EmailAttachment attachment) {
        if (_preparedAttachmentStreams.TryGetValue(attachment, out Stream? prepared)) {
            _preparedAttachmentStreams.Remove(attachment);
            prepared.Position = 0;
            return prepared;
        }
        return EmailAttachmentStreamScope.OpenRead(attachment);
    }

    private string EnsureTemporaryDirectory() {
        if (_temporaryDirectoryPath != null) return _temporaryDirectoryPath;
        _temporaryDirectoryPath = Path.Combine(Path.GetTempPath(),
            string.Concat("OfficeIMO.Email.Mime.", Guid.NewGuid().ToString("N")));
        EmailTemporaryStorage.CreatePrivateDirectory(_temporaryDirectoryPath);
        return _temporaryDirectoryPath;
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        foreach (Stream stream in _preparedAttachmentStreams.Values) stream.Dispose();
        _preparedAttachmentStreams.Clear();
        if (_temporaryDirectoryPath == null) return;
        try {
            if (Directory.Exists(_temporaryDirectoryPath)) Directory.Delete(_temporaryDirectoryPath, recursive: true);
        } catch {
            // Writer-owned staging cleanup is best effort and must not hide the serialization result.
        }
    }

    private sealed class AttachmentReferenceComparer : IEqualityComparer<EmailAttachment> {
        internal static AttachmentReferenceComparer Instance { get; } = new AttachmentReferenceComparer();
        public bool Equals(EmailAttachment? x, EmailAttachment? y) => ReferenceEquals(x, y);
        public int GetHashCode(EmailAttachment value) =>
            System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value);
    }
}
