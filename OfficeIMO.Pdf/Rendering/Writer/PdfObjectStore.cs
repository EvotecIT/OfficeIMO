using System.Collections;
using System.Security.Cryptography;

namespace OfficeIMO.Pdf;

/// <summary>Mutable serialized-object table that spills object bodies to a temporary file after a bounded memory budget.</summary>
internal sealed class PdfObjectStore : IList<byte[]>, IReadOnlyList<byte[]>, IDisposable {
    internal const long DefaultMemoryLimitBytes = 16L * 1024L * 1024L;

    private readonly List<Entry> _entries = new List<Entry>();
    private readonly long _memoryLimitBytes;
    private long _memoryBytes;
    private long _peakMemoryBytes;
    private FileStream? _spillStream;
    private string? _spillPath;
    private byte[]? _copyBuffer;
    private bool _disposed;

    internal PdfObjectStore(long memoryLimitBytes = DefaultMemoryLimitBytes) {
        if (memoryLimitBytes < 0L) throw new ArgumentOutOfRangeException(nameof(memoryLimitBytes), memoryLimitBytes, "PDF object-buffer memory limit cannot be negative.");
        _memoryLimitBytes = memoryLimitBytes;
    }

    internal bool IsSpilled => _spillStream != null;
    internal string? SpillPath => _spillPath;
    internal long RetainedMemoryBytes => _memoryBytes;
    internal long PeakRetainedMemoryBytes => _peakMemoryBytes;
    public int Count => _entries.Count;
    public bool IsReadOnly => false;

    public byte[] this[int index] {
        get {
            ThrowIfDisposed();
            Entry entry = _entries[index];
            if (entry.Bytes != null) return entry.Bytes;
            FileStream stream = _spillStream ?? throw new InvalidOperationException("PDF object spill storage is unavailable.");
            var bytes = new byte[entry.Length];
            stream.Position = entry.Offset;
            int read = 0;
            while (read < bytes.Length) {
                int count = stream.Read(bytes, read, bytes.Length - read);
                if (count == 0) throw new EndOfStreamException("PDF object spill storage ended unexpectedly.");
                read += count;
            }
            return bytes;
        }
        set {
            ThrowIfDisposed();
            Guard.NotNull(value, nameof(value));
            Entry previous = _entries[index];
            if (_spillStream == null && _memoryBytes - previous.Length + value.LongLength <= _memoryLimitBytes) {
                _entries[index] = Entry.InMemory(value);
                _memoryBytes = _memoryBytes - previous.Length + value.LongLength;
                _peakMemoryBytes = Math.Max(_peakMemoryBytes, _memoryBytes);
                return;
            }

            EnsureSpillStorage();
            _entries[index] = AppendToSpill(value);
        }
    }

    public void Add(byte[] item) {
        ThrowIfDisposed();
        Guard.NotNull(item, nameof(item));
        if (_spillStream == null && _memoryBytes + item.LongLength <= _memoryLimitBytes) {
            _entries.Add(Entry.InMemory(item));
            _memoryBytes += item.LongLength;
            _peakMemoryBytes = Math.Max(_peakMemoryBytes, _memoryBytes);
            return;
        }

        EnsureSpillStorage();
        _entries.Add(AppendToSpill(item));
    }

    internal void AddSegments(params byte[][] segments) {
        ThrowIfDisposed();
        Guard.NotNull(segments, nameof(segments));

        long totalLength = 0L;
        for (int index = 0; index < segments.Length; index++) {
            Guard.NotNull(segments[index], nameof(segments));
            totalLength += segments[index].LongLength;
        }
        if (totalLength > int.MaxValue) throw new InvalidOperationException("A serialized PDF object cannot exceed the supported two-gigabyte object size.");

        if (_spillStream == null && _memoryBytes + totalLength <= _memoryLimitBytes) {
            var bytes = new byte[(int)totalLength];
            int offset = 0;
            for (int index = 0; index < segments.Length; index++) {
                byte[] segment = segments[index];
                Buffer.BlockCopy(segment, 0, bytes, offset, segment.Length);
                offset += segment.Length;
            }
            _entries.Add(Entry.InMemory(bytes));
            _memoryBytes += totalLength;
            _peakMemoryBytes = Math.Max(_peakMemoryBytes, _memoryBytes);
            return;
        }

        EnsureSpillStorage();
        _entries.Add(AppendSegmentsToSpill(segments, (int)totalLength));
    }

    internal long GetLength(int index) {
        ThrowIfDisposed();
        return _entries[index].Length;
    }

    internal void CopyTo(int index, Stream destination, HashAlgorithm? hash = null) {
        ThrowIfDisposed();
        Guard.NotNull(destination, nameof(destination));
        if (!destination.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(destination));

        Entry entry = _entries[index];
        if (entry.Bytes != null) {
            hash?.TransformBlock(entry.Bytes, 0, entry.Bytes.Length, entry.Bytes, 0);
            destination.Write(entry.Bytes, 0, entry.Bytes.Length);
            return;
        }

        FileStream source = _spillStream ?? throw new InvalidOperationException("PDF object spill storage is unavailable.");
        source.Position = entry.Offset;
        byte[] buffer = _copyBuffer ??= new byte[81920];
        int remaining = entry.Length;
        while (remaining > 0) {
            int read = source.Read(buffer, 0, Math.Min(buffer.Length, remaining));
            if (read == 0) throw new EndOfStreamException("PDF object spill storage ended unexpectedly.");
            hash?.TransformBlock(buffer, 0, read, buffer, 0);
            destination.Write(buffer, 0, read);
            remaining -= read;
        }
    }

    public void Clear() {
        ThrowIfDisposed();
        _entries.Clear();
        _memoryBytes = 0L;
        _peakMemoryBytes = 0L;
        if (_spillStream != null) {
            _spillStream.SetLength(0L);
            _spillStream.Position = 0L;
        }
    }

    public bool Contains(byte[] item) => IndexOf(item) >= 0;

    public void CopyTo(byte[][] array, int arrayIndex) {
        ThrowIfDisposed();
        Guard.NotNull(array, nameof(array));
        if (arrayIndex < 0 || arrayIndex > array.Length || array.Length - arrayIndex < Count) throw new ArgumentOutOfRangeException(nameof(arrayIndex));
        for (int index = 0; index < Count; index++) array[arrayIndex + index] = this[index];
    }

    public IEnumerator<byte[]> GetEnumerator() {
        ThrowIfDisposed();
        for (int index = 0; index < Count; index++) yield return this[index];
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public int IndexOf(byte[] item) {
        ThrowIfDisposed();
        if (item == null) return -1;
        for (int index = 0; index < Count; index++) {
            byte[] candidate = this[index];
            if (ReferenceEquals(candidate, item) || candidate.SequenceEqual(item)) return index;
        }
        return -1;
    }

    public void Insert(int index, byte[] item) => throw new NotSupportedException("PDF object numbers are append-only while a document is assembled.");
    public bool Remove(byte[] item) => throw new NotSupportedException("PDF objects cannot be removed while a document is assembled.");
    public void RemoveAt(int index) => throw new NotSupportedException("PDF objects cannot be removed while a document is assembled.");

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _entries.Clear();
        _memoryBytes = 0L;
        _peakMemoryBytes = 0L;
        _copyBuffer = null;
        _spillStream?.Dispose();
        _spillStream = null;
        if (_spillPath != null) {
            try { File.Delete(_spillPath); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            _spillPath = null;
        }
    }

    private void EnsureSpillStorage() {
        if (_spillStream != null) return;
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO.Pdf-" + Guid.NewGuid().ToString("N") + ".objects");
        var stream = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.Read, 81920, FileOptions.None);
        _spillPath = path;
        _spillStream = stream;
        for (int index = 0; index < _entries.Count; index++) {
            Entry entry = _entries[index];
            if (entry.Bytes == null) continue;
            _entries[index] = AppendToSpill(entry.Bytes);
        }
        _memoryBytes = 0L;
    }

    private Entry AppendToSpill(byte[] bytes) {
        FileStream stream = _spillStream ?? throw new InvalidOperationException("PDF object spill storage is unavailable.");
        stream.Position = stream.Length;
        long offset = stream.Position;
        stream.Write(bytes, 0, bytes.Length);
        return Entry.Spilled(offset, bytes.Length);
    }

    private Entry AppendSegmentsToSpill(byte[][] segments, int totalLength) {
        FileStream stream = _spillStream ?? throw new InvalidOperationException("PDF object spill storage is unavailable.");
        stream.Position = stream.Length;
        long offset = stream.Position;
        for (int index = 0; index < segments.Length; index++) {
            byte[] segment = segments[index];
            stream.Write(segment, 0, segment.Length);
        }
        return Entry.Spilled(offset, totalLength);
    }

    #pragma warning disable CA1513 // Newer helper is unavailable on netstandard2.0 and net472.
    private void ThrowIfDisposed() {
        if (_disposed) throw new ObjectDisposedException(nameof(PdfObjectStore));
    }
    #pragma warning restore CA1513

    private readonly struct Entry {
        private Entry(byte[]? bytes, long offset, int length) {
            Bytes = bytes;
            Offset = offset;
            Length = length;
        }

        internal byte[]? Bytes { get; }
        internal long Offset { get; }
        internal int Length { get; }
        internal static Entry InMemory(byte[] bytes) => new Entry(bytes, 0L, bytes.Length);
        internal static Entry Spilled(long offset, int length) => new Entry(null, offset, length);
    }
}
