namespace OfficeIMO.Pdf;

/// <summary>Bounded page-content storage that spills completed page streams to a temporary file.</summary>
internal sealed class PdfPageContentStore : IDisposable {
    internal const long DefaultMemoryLimitBytes = 4L * 1024L * 1024L;

    private readonly List<Entry> _entries = new List<Entry>();
    private readonly long _memoryLimitBytes;
    private long _memoryBytes;
    private long _peakMemoryBytes;
    private FileStream? _spillStream;
    private string? _spillPath;
    private bool _disposed;

    internal PdfPageContentStore(long memoryLimitBytes = DefaultMemoryLimitBytes) {
#pragma warning disable CA1512 // Cross-target guard supports netstandard2.0 and net472.
        if (memoryLimitBytes < 0L) throw new ArgumentOutOfRangeException(nameof(memoryLimitBytes));
#pragma warning restore CA1512
        _memoryLimitBytes = memoryLimitBytes;
    }

    internal bool IsSpilled => _spillStream != null;
    internal string? SpillPath => _spillPath;
    internal long RetainedMemoryBytes => _memoryBytes;
    internal long PeakRetainedMemoryBytes => _peakMemoryBytes;

    internal PdfPageContentHandle Store(string content) {
        ThrowIfDisposed();
        byte[] bytes = PdfEncoding.Latin1GetBytes(content ?? string.Empty);
        int index = _entries.Count;
        if (_spillStream == null && _memoryBytes + bytes.LongLength <= _memoryLimitBytes) {
            _entries.Add(Entry.InMemory(bytes));
            _memoryBytes += bytes.LongLength;
            _peakMemoryBytes = Math.Max(_peakMemoryBytes, _memoryBytes);
        } else {
            EnsureSpillStorage();
            _entries.Add(AppendToSpill(bytes));
        }
        return new PdfPageContentHandle(index);
    }

    internal string Read(PdfPageContentHandle handle) {
        ThrowIfDisposed();
        if (handle.Index < 0 || handle.Index >= _entries.Count) throw new ArgumentOutOfRangeException(nameof(handle));
        Entry entry = _entries[handle.Index];
        if (entry.Bytes != null) return PdfEncoding.Latin1GetString(entry.Bytes);
        var bytes = new byte[entry.Length];
        FileStream stream = _spillStream ?? throw new InvalidOperationException("Page-content spill storage is unavailable.");
        stream.Position = entry.Offset;
        ReadExactly(stream, bytes);
        return PdfEncoding.Latin1GetString(bytes);
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _entries.Clear();
        _memoryBytes = 0L;
        _peakMemoryBytes = 0L;
        _spillStream?.Dispose();
        _spillStream = null;
        if (_spillPath != null) {
            try { File.Delete(_spillPath); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            _spillPath = null;
        }
    }

    private void EnsureSpillStorage() {
        if (_spillStream != null) return;
        _spillStream = PdfTemporaryFile.Create(".pages", FileOptions.SequentialScan, out string path);
        _spillPath = path;
        for (int i = 0; i < _entries.Count; i++) {
            Entry entry = _entries[i];
            if (entry.Bytes != null) _entries[i] = AppendToSpill(entry.Bytes);
        }
        _memoryBytes = 0L;
    }

    private Entry AppendToSpill(byte[] bytes) {
        FileStream stream = _spillStream ?? throw new InvalidOperationException("Page-content spill storage is unavailable.");
        stream.Position = stream.Length;
        long offset = stream.Position;
        stream.Write(bytes, 0, bytes.Length);
        return Entry.Spilled(offset, bytes.Length);
    }

    private static void ReadExactly(Stream stream, byte[] bytes) {
        int offset = 0;
        while (offset < bytes.Length) {
            int count = stream.Read(bytes, offset, bytes.Length - offset);
            if (count == 0) throw new EndOfStreamException("Page-content spill storage ended unexpectedly.");
            offset += count;
        }
    }

    private void ThrowIfDisposed() {
#pragma warning disable CA1513 // Newer helper is unavailable on netstandard2.0 and net472.
        if (_disposed) throw new ObjectDisposedException(nameof(PdfPageContentStore));
#pragma warning restore CA1513
    }

    private readonly struct Entry {
        private Entry(byte[]? bytes, long offset, int length) { Bytes = bytes; Offset = offset; Length = length; }
        internal byte[]? Bytes { get; }
        internal long Offset { get; }
        internal int Length { get; }
        internal static Entry InMemory(byte[] bytes) => new Entry(bytes, 0L, bytes.Length);
        internal static Entry Spilled(long offset, int length) => new Entry(null, offset, length);
    }
}

internal readonly struct PdfPageContentHandle {
    internal PdfPageContentHandle(int index) { Index = index; }
    internal int Index { get; }
}
