namespace OfficeIMO.Email;

/// <summary>Stages asynchronous attachment sources to bounded, reopenable writer-owned files.</summary>
internal sealed class EmailAttachmentStaging : IDisposable {
    private readonly Dictionary<EmailAttachment, IEmailContentSource> _sources =
        new Dictionary<EmailAttachment, IEmailContentSource>(ReferenceComparer.Instance);
    private string? _directoryPath;
    private bool _disposed;

    private EmailAttachmentStaging() { }

    internal static EmailAttachmentStaging CreateEmpty() => new EmailAttachmentStaging();

    internal static async Task<EmailAttachmentStaging> CreateAsync(EmailDocument document,
        long maximumBytes, CancellationToken cancellationToken, Action<int>? bytesRead = null,
        int maxDepth = 16, int maxAttachments = int.MaxValue, bool includeTaskPayload = false) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        var staging = new EmailAttachmentStaging();
        try {
            var visited = new HashSet<EmailDocument>(DocumentReferenceComparer.Instance);
            long totalBytes = 0;
            await staging.StageDocumentAsync(document, visited, maximumBytes,
                new MutableLong(totalBytes), cancellationToken, bytesRead, 0, maxDepth,
                new MutableLong(0), maxAttachments, includeTaskPayload).ConfigureAwait(false);
            return staging;
        } catch {
            staging.Dispose();
            throw;
        }
    }

    internal IDisposable EnterScope() => EmailAttachmentStreamScope.Begin(_sources);

    private async Task StageDocumentAsync(EmailDocument document, ISet<EmailDocument> visited,
        long maximumBytes, MutableLong totalBytes, CancellationToken cancellationToken, Action<int>? bytesRead,
        int depth, int maxDepth, MutableLong attachmentCount, int maxAttachments, bool includeTaskPayload) {
        if (depth > maxDepth) throw new EmailLimitExceededException(nameof(EmailWriterOptions.MaxNestedMessageDepth), depth, maxDepth);
        if (!visited.Add(document)) return;
        foreach (EmailAttachment attachment in document.Attachments) {
            cancellationToken.ThrowIfCancellationRequested();
            if (++attachmentCount.Value > maxAttachments) throw new EmailLimitExceededException("MaxAttachments", attachmentCount.Value, maxAttachments);
            if (attachment.EmbeddedDocument != null) {
                await StageDocumentAsync(attachment.EmbeddedDocument, visited, maximumBytes,
                    totalBytes, cancellationToken, bytesRead, depth + 1, maxDepth, attachmentCount, maxAttachments, includeTaskPayload).ConfigureAwait(false);
                continue;
            }
            if (attachment.Content != null || attachment.ContentSource == null || _sources.ContainsKey(attachment)) {
                continue;
            }

            long? declaredLength = attachment.ContentSource.Length;
            if (declaredLength.HasValue && declaredLength.Value > maximumBytes - totalBytes.Value) {
                throw new EmailLimitExceededException(nameof(EmailWriterOptions.MaxOutputBytes),
                    checked(totalBytes.Value + declaredLength.Value), maximumBytes);
            }

            string directory = EnsureDirectory();
            string path = Path.Combine(directory,
                string.Concat(_sources.Count.ToString("D8", CultureInfo.InvariantCulture), ".content"));
            long stagedLength = 0;
            using (Stream input = await attachment.OpenContentStreamAsync(cancellationToken).ConfigureAwait(false))
            using (var output = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read,
                       81920, FileOptions.Asynchronous | FileOptions.SequentialScan)) {
                var buffer = new byte[81920];
                while (true) {
                    cancellationToken.ThrowIfCancellationRequested();
                    long remaining = maximumBytes - totalBytes.Value - stagedLength;
                    int requested = remaining >= buffer.Length ? buffer.Length : (int)remaining + 1;
                    int read = await input.ReadAsync(buffer, 0, requested, cancellationToken).ConfigureAwait(false);
                    bytesRead?.Invoke(read);
                    cancellationToken.ThrowIfCancellationRequested();
                    if (read == 0) break;
                    stagedLength = checked(stagedLength + read);
                    long aggregate = checked(totalBytes.Value + stagedLength);
                    if (aggregate > maximumBytes) {
                        throw new EmailLimitExceededException(nameof(EmailWriterOptions.MaxOutputBytes),
                            aggregate, maximumBytes);
                    }
                    await output.WriteAsync(buffer, 0, read, cancellationToken).ConfigureAwait(false);
                }
                await output.FlushAsync(cancellationToken).ConfigureAwait(false);
            }
            totalBytes.Value = checked(totalBytes.Value + stagedLength);
            _sources.Add(attachment, new StagedFileContentSource(path, stagedLength));
        }
        if (includeTaskPayload && OutlookTaskCommunicationAttachmentProjection.GetEmbeddedTaskForWriting(document) is EmailDocument task &&
            !visited.Contains(task)) {
            if (++attachmentCount.Value > maxAttachments) throw new EmailLimitExceededException("MaxAttachments", attachmentCount.Value, maxAttachments);
            await StageDocumentAsync(task, visited, maximumBytes, totalBytes, cancellationToken, bytesRead,
                depth + 1, maxDepth, attachmentCount, maxAttachments, includeTaskPayload).ConfigureAwait(false);
        }
    }

    private string EnsureDirectory() {
        if (_directoryPath != null) return _directoryPath;
        _directoryPath = Path.Combine(Path.GetTempPath(),
            string.Concat("OfficeIMO.Email.Write.", Guid.NewGuid().ToString("N")));
        EmailTemporaryStorage.CreatePrivateDirectory(_directoryPath);
        return _directoryPath;
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        if (_directoryPath == null) return;
        try {
            if (Directory.Exists(_directoryPath)) Directory.Delete(_directoryPath, recursive: true);
        } catch {
            // Writer-owned staging cleanup is best effort and must not hide the serialization result.
        }
    }

    private sealed class MutableLong {
        internal MutableLong(long value) { Value = value; }
        internal long Value { get; set; }
    }

    private sealed class StagedFileContentSource : IEmailContentSource {
        private readonly string _path;
        internal StagedFileContentSource(string path, long length) { _path = path; Length = length; }
        public long? Length { get; }
        public Stream OpenRead() => new FileStream(_path, FileMode.Open, FileAccess.Read,
            FileShare.Read, 81920, FileOptions.SequentialScan);
        public Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult<Stream>(new FileStream(_path, FileMode.Open, FileAccess.Read,
                FileShare.Read, 81920, FileOptions.Asynchronous | FileOptions.SequentialScan));
        }
    }

    private sealed class ReferenceComparer : IEqualityComparer<EmailAttachment> {
        internal static ReferenceComparer Instance { get; } = new ReferenceComparer();
        public bool Equals(EmailAttachment? x, EmailAttachment? y) => ReferenceEquals(x, y);
        public int GetHashCode(EmailAttachment value) =>
            System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value);
    }

    private sealed class DocumentReferenceComparer : IEqualityComparer<EmailDocument> {
        internal static DocumentReferenceComparer Instance { get; } = new DocumentReferenceComparer();
        public bool Equals(EmailDocument? x, EmailDocument? y) => ReferenceEquals(x, y);
        public int GetHashCode(EmailDocument value) =>
            System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value);
    }
}
