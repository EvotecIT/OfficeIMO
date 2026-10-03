using Avalonia.Platform.Storage;
using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;

namespace OfficeIMO.Studio.Infrastructure;

internal sealed partial class StudioStorageAccess {
    private static string OpenedIdentity(string location, Stream stream) {
        string? path = OfficeStorageIdentity.GetLocalPath(location);
        if (path is null) return OfficeStorageIdentity.Normalize(location);
        while (stream is PermissionStream scoped) stream = scoped.Inner;
        string identity = stream is FileStream file
            ? OfficePathIdentity.GetPhysicalIdentityKey(path, file.SafeFileHandle)
            : OfficePathIdentity.GetPhysicalIdentityKey(path);
        if (OfficePathIdentity.GetPhysicalIdentityKey(path) != identity) {
            throw new IOException("The document changed while it was being opened. Select it again to read the current file.");
        }
        return identity;
    }

    private void RefreshNativePermission(string location) {
        if (!OfficeMacFilePermission.IsSandboxed || OfficeStorageIdentity.GetLocalPath(location) is not { } path) return;
        string key = OfficeStorageIdentity.Normalize(location);
        StudioStorageReference? reference;
        lock (_sync) _references.TryGetValue(key, out reference);
        if (reference is null || OfficeMacFilePermission.IsNativeBookmark(reference.Bookmark)) return;
        // This runs only while the selected provider's read stream holds access,
        // including verification reads after a newly selected output is created.
        string bookmark = OfficeMacFilePermission.CreateBookmark(path);
        lock (_sync) {
            ObjectDisposedException.ThrowIf(_disposed, this);
            _references[key] = reference with { Bookmark = bookmark };
        }
    }

    private OfficeMacFilePermission? OpenNativePermission(string location) {
        StudioStorageReference? reference;
        lock (_sync) _references.TryGetValue(OfficeStorageIdentity.Normalize(location), out reference);
        return reference?.Bookmark is { } bookmark && OfficeMacFilePermission.IsNativeBookmark(bookmark)
            ? OfficeMacFilePermission.Open(bookmark, OfficeStorageIdentity.GetLocalPath(location)
                ?? throw new IOException("Native permissions require a local file.")) : null;
    }

    private async Task<Stream> OpenWriteAsync(string location, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        OfficeMacFilePermission? permission = OpenNativePermission(location);
        try {
            IStorageFile file = await ResolveAsync(location, token).ConfigureAwait(false)
                ?? throw new IOException("The storage provider is unavailable. Select the destination again.");
            Stream stream = await file.OpenWriteAsync().ConfigureAwait(false);
            try { token.ThrowIfCancellationRequested(); }
            catch { await stream.DisposeAsync(); throw; }
            return permission is null ? stream : new PermissionStream(stream, permission);
        } catch { permission?.Dispose(); throw; }
    }

    private sealed class PermissionStream(Stream inner, OfficeMacFilePermission permission) : Stream {
        internal Stream Inner => inner;
        public override bool CanRead => inner.CanRead;
        public override bool CanWrite => inner.CanWrite;
        public override bool CanSeek => inner.CanSeek;
        public override long Length => inner.Length;
        public override long Position { get => inner.Position; set => inner.Position = value; }
        public override void Flush() => inner.Flush();
        public override Task FlushAsync(CancellationToken token) => inner.FlushAsync(token);
        public override int Read(byte[] buffer, int offset, int count) => inner.Read(buffer, offset, count);
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken token) => inner.ReadAsync(buffer, offset, count, token);
        public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken token = default) => inner.ReadAsync(buffer, token);
        public override void Write(byte[] buffer, int offset, int count) => inner.Write(buffer, offset, count);
        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken token) => inner.WriteAsync(buffer, offset, count, token);
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken token = default) => inner.WriteAsync(buffer, token);
        public override long Seek(long offset, SeekOrigin origin) => inner.Seek(offset, origin);
        public override void SetLength(long value) => inner.SetLength(value);
        protected override void Dispose(bool disposing) {
            if (disposing) { try { inner.Dispose(); } finally { permission.Dispose(); } }
            base.Dispose(disposing);
        }
        public override async ValueTask DisposeAsync() {
            try { await inner.DisposeAsync(); } finally { permission.Dispose(); }
            GC.SuppressFinalize(this);
        }
    }
}
