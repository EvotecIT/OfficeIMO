namespace OfficeIMO.Email;

/// <summary>Owns exact temporary files used by one streaming read result.</summary>
internal sealed class EmailReadWorkspace : IDisposable {
    private readonly string _directoryPath;
    private readonly Dictionary<string, IEmailContentSource> _sources =
        new Dictionary<string, IEmailContentSource>(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<WorkspaceReadStream> _openReaders = new HashSet<WorkspaceReadStream>();
    private bool _disposed;

    internal EmailReadWorkspace() {
        _directoryPath = Path.Combine(Path.GetTempPath(),
            string.Concat("OfficeIMO.Email.Read.", Guid.NewGuid().ToString("N")));
        EmailTemporaryStorage.CreatePrivateDirectory(_directoryPath);
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
        foreach (WorkspaceReadStream stream in _openReaders.ToArray()) stream.Dispose();
        DeleteDirectory();
        GC.SuppressFinalize(this);
    }

    ~EmailReadWorkspace() => DeleteDirectory();

    private Stream OpenContent(string path, bool asynchronous) {
        EnsureAlive();
        var stream = new WorkspaceReadStream(this, new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read,
            81920, FileOptions.SequentialScan | (asynchronous ? FileOptions.Asynchronous : FileOptions.None)));
        _openReaders.Add(stream);
        return stream;
    }

    private sealed class WorkspaceReadStream : Stream {
        private readonly EmailReadWorkspace _owner;
        private readonly Stream _inner;
        internal WorkspaceReadStream(EmailReadWorkspace owner, Stream inner) { _owner = owner; _inner = inner; }
        public override bool CanRead => !_owner._disposed && _inner.CanRead;
        public override bool CanSeek => !_owner._disposed && _inner.CanSeek;
        public override bool CanWrite => false;
        public override long Length { get { _owner.EnsureAlive(); return _inner.Length; } }
        public override long Position {
            get { _owner.EnsureAlive(); return _inner.Position; }
            set { _owner.EnsureAlive(); _inner.Position = value; }
        }
        public override int Read(byte[] buffer, int offset, int count) {
            _owner.EnsureAlive();
            return _inner.Read(buffer, offset, count);
        }
        public override int ReadByte() { _owner.EnsureAlive(); return _inner.ReadByte(); }
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
            _owner.EnsureAlive();
            return _inner.ReadAsync(buffer, offset, count, cancellationToken);
        }
        public override long Seek(long offset, SeekOrigin origin) { _owner.EnsureAlive(); return _inner.Seek(offset, origin); }
        public override void Flush() { }
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) {
            if (disposing) {
                _inner.Dispose();
                _owner._openReaders.Remove(this);
            }
            base.Dispose(disposing);
        }
    }

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
            return _workspace.OpenContent(_path, asynchronous: false);
        }
        public Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(_workspace.OpenContent(_path, asynchronous: true));
        }
    }
}
