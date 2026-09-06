namespace OfficeIMO.Workflows.Tests;

/// <summary>Models a provider whose filesystem name is accessible only during its returned stream's lifetime.</summary>
internal sealed class ScopedWorkflowFile {
    internal ScopedWorkflowFile(string root, string name, byte[] bytes) {
        Path = System.IO.Path.Combine(root, name);
        BackingPath = Path + ".outside-scope";
        File.WriteAllBytes(BackingPath, bytes);
    }
    internal string Path { get; }
    internal string BackingPath { get; }
    internal int Opens { get; private set; }
    internal int Closes { get; private set; }
    internal int Writes { get; private set; }
    private int _activeScopes;
    internal Task<Stream> OpenRead(CancellationToken token) => Open(false, token);
    internal Task<Stream> OpenWrite(CancellationToken token) => Open(true, token);
    private Task<Stream> Open(bool write, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (_activeScopes == 0) File.Move(BackingPath, Path);
        try {
            var stream = new ScopedStream(Path, write, () => {
                if (--_activeScopes == 0) File.Move(Path, BackingPath);
                Closes++;
            });
            _activeScopes++;
            Opens++;
            if (write) Writes++;
            return Task.FromResult<Stream>(stream);
        } catch {
            if (_activeScopes == 0) File.Move(Path, BackingPath);
            throw;
        }
    }

    private sealed class ScopedStream(string path, bool write, Action close)
        : FileStream(path, write ? FileMode.Create : FileMode.Open, write ? FileAccess.Write : FileAccess.Read, FileShare.Read) {
        private bool _closed;
        private void CloseScope() { if (!_closed) { _closed = true; close(); } }
        protected override void Dispose(bool disposing) { base.Dispose(disposing); if (disposing) CloseScope(); }
        public override async ValueTask DisposeAsync() { await base.DisposeAsync(); CloseScope(); }
    }
}
