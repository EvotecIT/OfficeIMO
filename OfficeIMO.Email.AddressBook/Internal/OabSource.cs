namespace OfficeIMO.Email.AddressBook;

internal sealed class OabSource {
    private readonly string? _filePath;
    private readonly Stream? _stream;
    private readonly long _restorePosition;

    private OabSource(string sourcePath, string sourceName, string filePath, long length) {
        SourcePath = sourcePath;
        SourceName = sourceName;
        _filePath = filePath;
        Length = length;
    }

    private OabSource(string sourceName, Stream stream, long baseOffset, long length) {
        SourcePath = sourceName;
        SourceName = sourceName;
        _stream = stream;
        BaseOffset = baseOffset;
        _restorePosition = baseOffset;
        Length = length;
    }

    internal string SourcePath { get; }
    internal string SourceName { get; }
    internal long BaseOffset { get; }
    internal long Length { get; }
    internal bool UsesCallerStream => _stream != null;

    internal static OabSource FromFile(string path) {
        string fullPath = Path.GetFullPath(path);
        var file = new FileInfo(fullPath);
        return new OabSource(fullPath, file.Name, fullPath, file.Length);
    }

    internal static OabSource FromStream(Stream stream, string sourceName) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("OAB stream must be readable.", nameof(stream));
        if (!stream.CanSeek) throw new ArgumentException("OAB stream must be seekable for lazy record access.", nameof(stream));
        if (sourceName == null) throw new ArgumentNullException(nameof(sourceName));
        long offset = stream.Position;
        return new OabSource(sourceName, stream, offset, checked(stream.Length - offset));
    }

    internal OabStreamLease OpenRead() {
        if (_filePath != null) {
            var stream = new FileStream(_filePath, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete, 64 * 1024, FileOptions.SequentialScan);
            return new OabStreamLease(stream, ownsStream: true, restorePosition: null);
        }
        if (_stream == null) throw new ObjectDisposedException(nameof(OabSource));
        long current = _stream.Position;
        _stream.Position = BaseOffset;
        return new OabStreamLease(_stream, ownsStream: false, restorePosition: current);
    }

    internal void RestoreCallerPosition() {
        if (_stream != null && _stream.CanSeek) _stream.Position = _restorePosition;
    }
}

internal sealed class OabStreamLease : IDisposable {
    private readonly bool _ownsStream;
    private readonly long? _restorePosition;
    private bool _disposed;

    internal OabStreamLease(Stream stream, bool ownsStream, long? restorePosition) {
        Stream = stream;
        _ownsStream = ownsStream;
        _restorePosition = restorePosition;
    }

    internal Stream Stream { get; }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        if (_ownsStream) {
            Stream.Dispose();
        } else if (_restorePosition.HasValue && Stream.CanSeek) {
            Stream.Position = _restorePosition.Value;
        }
    }
}
