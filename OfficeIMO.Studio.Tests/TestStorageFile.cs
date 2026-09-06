using System.Reflection;
using Avalonia.Platform.Storage;

namespace OfficeIMO.Studio.Tests;

/// <summary>A stream-only provider boundary for document lifecycle and publication tests.</summary>
internal sealed class TestStorageFile {
    internal IStorageBookmarkFile Item { get; }
    internal Uri Location { get; }
    internal string Name { get; }
    internal string Bookmark { get; }
    internal byte[] Bytes { get; set; }
    internal int Reads { get; private set; }
    internal int ClosedReads { get; private set; }
    internal int Writes { get; private set; }
    internal int Disposals { get; private set; }
    internal bool DenyRead { get; set; }
    internal bool FailWrite { get; set; }
    internal bool CorruptWrite { get; set; }
    internal Action? BeforeRead { get; set; }

    internal TestStorageFile(string location, byte[] bytes, string name = "Provider document.pdf", string? bookmark = null) {
        Location = new(location);
        Name = name;
        Bytes = bytes;
        Bookmark = bookmark ?? Guid.NewGuid().ToString("N");
        Item = DispatchProxy.Create<IStorageBookmarkFile, StorageProxy>();
        ((StorageProxy)(object)Item).Call = (method, _) => method switch {
            "get_Name" => Name,
            "get_Path" => Location,
            "get_CanBookmark" => true,
            "SaveBookmarkAsync" => Task.FromResult<string?>(Bookmark),
            "ReleaseBookmarkAsync" => Task.CompletedTask,
            "OpenReadAsync" => OpenRead(),
            "OpenWriteAsync" => OpenWrite(),
            "Dispose" => Release(),
            _ => throw new NotSupportedException(method)
        };
    }

    internal IStorageProvider CreateProvider(bool denyBookmark = false) {
        var provider = DispatchProxy.Create<IStorageProvider, StorageProxy>();
        ((StorageProxy)(object)provider).Call = (method, args) => method switch {
            "OpenFileBookmarkAsync" => Task.FromResult<IStorageBookmarkFile?>(!denyBookmark && (string?)args![0] == Bookmark ? Item : null),
            "TryGetFileFromPathAsync" => Task.FromResult<IStorageFile?>(Equals(args![0], Location) ? Item : null),
            _ => throw new NotSupportedException(method)
        };
        return provider;
    }

    private Task<Stream> OpenRead() {
        Reads++;
        BeforeRead?.Invoke();
        if (DenyRead) throw new UnauthorizedAccessException("Provider permission expired.");
        return Task.FromResult<Stream>(new ReadStream(Bytes.ToArray(), () => ClosedReads++));
    }

    private Task<Stream> OpenWrite() {
        Writes++;
        Bytes = [];
        return Task.FromResult<Stream>(new WriteStream(bytes => Bytes = CorruptWrite ? [.. bytes, 0x21] : bytes, FailWrite));
    }

    private object? Release() { Disposals++; return null; }

    private sealed class ReadStream(byte[] bytes, Action close) : MemoryStream(bytes) {
        private bool _closed;
        public override bool CanSeek => false;
        protected override void Dispose(bool disposing) {
            if (!_closed) { _closed = true; close(); }
            base.Dispose(disposing);
        }
    }

    private sealed class WriteStream(Action<byte[]> close, bool fail) : MemoryStream {
        private bool _closed;
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken token = default) {
            token.ThrowIfCancellationRequested();
            if (!fail) return base.WriteAsync(buffer, token);
            base.Write(buffer.Span[..Math.Min(16, buffer.Length)]);
            throw new IOException("Provider write interrupted.");
        }
        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken token) {
            token.ThrowIfCancellationRequested();
            if (!fail) return base.WriteAsync(buffer, offset, count, token);
            base.Write(buffer, offset, Math.Min(16, count));
            throw new IOException("Provider write interrupted.");
        }
        protected override void Dispose(bool disposing) {
            if (!_closed) { _closed = true; close(ToArray()); }
            base.Dispose(disposing);
        }
    }

    public class StorageProxy : DispatchProxy {
        internal Func<string, object?[]?, object?> Call { get; set; } = null!;
        protected override object? Invoke(MethodInfo? method, object?[]? args) => Call(method!.Name, args);
    }
}
