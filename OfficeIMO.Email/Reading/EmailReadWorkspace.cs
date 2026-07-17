namespace OfficeIMO.Email;

/// <summary>Owns exact temporary files used by one streaming read result.</summary>
internal sealed class EmailReadWorkspace : IDisposable {
    private readonly string _directoryPath;
    private readonly Dictionary<string, IEmailContentSource> _sources =
        new Dictionary<string, IEmailContentSource>(StringComparer.OrdinalIgnoreCase);
    private bool _disposed;

    internal EmailReadWorkspace() {
        _directoryPath = Path.Combine(Path.GetTempPath(),
            string.Concat("OfficeIMO.Email.Read.", Guid.NewGuid().ToString("N")));
        Directory.CreateDirectory(_directoryPath);
    }

    internal Stream OpenExternalDestination(string logicalPath, long length) {
        EnsureAlive();
        string path = Path.Combine(_directoryPath,
            string.Concat("content-", _sources.Count.ToString("D8", CultureInfo.InvariantCulture), ".bin"));
        var source = new WorkspaceContentSource(this, path, length);
        _sources.Add(logicalPath, source);
        return new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read,
            81920, FileOptions.SequentialScan);
    }

    internal IReadOnlyDictionary<string, IEmailContentSource> GetSources() => _sources;

    internal bool HasContent => _sources.Count > 0;

    internal string CreateInputPath() {
        EnsureAlive();
        return Path.Combine(_directoryPath, "input.artifact");
    }

    internal string CreateContentPath() {
        EnsureAlive();
        return Path.Combine(_directoryPath,
            string.Concat("content-", _sources.Count.ToString("D8", CultureInfo.InvariantCulture), "-",
                Guid.NewGuid().ToString("N"), ".bin"));
    }

    internal IEmailContentSource RegisterContent(string logicalPath, string path, long length) {
        EnsureAlive();
        var source = new WorkspaceContentSource(this, path, length);
        _sources.Add(logicalPath, source);
        return source;
    }

    internal void EnsureAlive() {
        if (_disposed) throw new ObjectDisposedException(nameof(EmailReadResult));
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        DeleteDirectory();
        GC.SuppressFinalize(this);
    }

    ~EmailReadWorkspace() => DeleteDirectory();

    private void DeleteDirectory() {
        try {
            if (Directory.Exists(_directoryPath)) Directory.Delete(_directoryPath, recursive: true);
        } catch {
            // A finalizer or explicit result disposal must not throw during best-effort temporary cleanup.
        }
    }

    private sealed class WorkspaceContentSource : IEmailContentSource {
        private readonly EmailReadWorkspace _workspace;
        private readonly string _path;
        internal WorkspaceContentSource(EmailReadWorkspace workspace, string path, long length) {
            _workspace = workspace;
            _path = path;
            Length = length;
        }
        public long? Length { get; }
        public Stream OpenRead() {
            _workspace.EnsureAlive();
            return new FileStream(_path, FileMode.Open, FileAccess.Read, FileShare.Read,
                81920, FileOptions.SequentialScan);
        }
        public Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            _workspace.EnsureAlive();
            return Task.FromResult<Stream>(new FileStream(_path, FileMode.Open, FileAccess.Read, FileShare.Read,
                81920, FileOptions.Asynchronous | FileOptions.SequentialScan));
        }
    }
}
