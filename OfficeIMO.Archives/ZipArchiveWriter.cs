using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Archives;

/// <summary>The compression method for a written ZIP member.</summary>
public enum ZipCompressionMethod {
    /// <summary>Uncompressed member bytes (method 0).</summary>
    Stored = 0,
    /// <summary>Deflate-compressed member bytes (method 8).</summary>
    Deflate = 8
}

/// <summary>Settings for <see cref="ZipArchiveWriter"/>.</summary>
public sealed class ZipArchiveWriterOptions : ArchiveWriterOptions {
    /// <summary>The default member compression. Empty files are always stored.</summary>
    public ZipCompressionMethod Compression { get; set; } = ZipCompressionMethod.Deflate;
    /// <summary>When true, members are written straight to the destination and followed by data descriptors
    /// (general purpose bit 3) instead of being staged first. Use it for nonseekable destinations where temporary
    /// storage is unwanted. Members of unknown length are then limited to 4 GiB.</summary>
    public bool UseDataDescriptors { get; set; }
}

/// <summary>Writes ZIP archives (PKWARE APPNOTE 6.3.10) with stored or Deflate members, UTF-8 names and ZIP64 when sizes
/// or offsets require it. With the default options the output depends only on the entries written: timestamps are fixed,
/// attributes are constant, no varying extra fields are written and Deflate output is the same on every platform.</summary>
/// <remarks>Each member is compressed into a staging buffer (memory, then a private temporary file above 8 MiB) so its
/// local header carries exact sizes and checksums. The destination needs to be writable but not seekable.
/// Member names pass the same path checks as extraction.</remarks>
public sealed class ZipArchiveWriter : IDisposable {
    private const uint Max32 = 0xFFFFFFFF;
    private const ushort MadeBy = (3 << 8) | 63;
    private const uint FileAttributes = 0x81A40000; // Unix regular file, rw-r--r--
    private const uint DirectoryAttributes = 0x41ED0010; // Unix directory, rwxr-xr-x, plus the MS-DOS directory bit
    private static readonly DateTimeOffset DosEpoch = new(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);

    private readonly CountingStream _sink;
    private readonly Stream _destination;
    private readonly bool _leaveOpen;
    private readonly ZipArchiveWriterOptions _options;
    private readonly ArchivePaths _paths;
    private readonly List<CentralRecord> _records = new();
    private EntryStream? _open;
    private bool _finished;
    private bool _faulted;
    private bool _disposed;

    /// <summary>Starts an archive at the destination's current position.</summary>
    public ZipArchiveWriter(Stream destination, ZipArchiveWriterOptions? options = null, bool leaveOpen = false) {
        ArgumentNullException.ThrowIfNull(destination);
        if (!destination.CanWrite) throw new ArgumentException("A writable destination stream is required.", nameof(destination));
        _destination = destination; _leaveOpen = leaveOpen;
        _options = options ?? new ZipArchiveWriterOptions();
        if (_options.Compression != ZipCompressionMethod.Stored && _options.Compression != ZipCompressionMethod.Deflate)
            throw new ArgumentOutOfRangeException(nameof(options), "Unsupported ZIP compression method.");
        _paths = _options.CreatePaths();
        _sink = new CountingStream(destination);
    }

    /// <summary>Number of members written so far.</summary>
    public int Count => _records.Count;

    /// <summary>Opens a write stream for one file member. Dispose it to complete the member before adding another.</summary>
    public Stream CreateEntry(string path, DateTimeOffset? lastWriteTime = null, ZipCompressionMethod? compression = null) =>
        Begin(path, lastWriteTime, compression, knownLength: null);

    /// <summary>Adds one file member from the stream's current position to its end. The stream is left open.</summary>
    public void AddFile(string path, Stream content, DateTimeOffset? lastWriteTime = null, ZipCompressionMethod? compression = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(content);
        long? length = content.CanSeek ? content.Length - content.Position : null;
        EntryStream entry = Begin(path, lastWriteTime, compression, length);
        try {
            var buffer = new byte[81920];
            int read;
            while ((read = content.Read(buffer, 0, buffer.Length)) > 0) {
                cancellationToken.ThrowIfCancellationRequested();
                entry.Write(buffer, 0, read);
            }
            cancellationToken.ThrowIfCancellationRequested();
        } catch { Abort(entry); throw; }
        entry.Dispose();
    }

    /// <summary>Adds one file member from memory.</summary>
    public void AddFile(string path, byte[] content, DateTimeOffset? lastWriteTime = null, ZipCompressionMethod? compression = null) {
        ArgumentNullException.ThrowIfNull(content);
        EntryStream entry = Begin(path, lastWriteTime, compression, content.LongLength);
        try { entry.Write(content, 0, content.Length); } catch { Abort(entry); throw; }
        entry.Dispose();
    }

    /// <summary>Adds an explicit directory member, written with a trailing '/'.</summary>
    public void AddDirectory(string path, DateTimeOffset? lastWriteTime = null) {
        ThrowIfUnavailable();
        ArgumentNullException.ThrowIfNull(path);
        string normalized = path.Replace('\\', '/').TrimEnd('/');
        _paths.AddDirectory(normalized);
        byte[] name = Encoding.UTF8.GetBytes(normalized + "/");
        var record = new CentralRecord(name, Flags(name, false), 0, Dos(lastWriteTime), true, _sink.Position);
        record.Complete(0, 0, 0);
        WriteLocalHeader(record);
        _records.Add(record);
    }

    /// <summary>Adds one described entry.</summary>
    public void Add(ArchiveWriteEntry entry, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(entry);
        if (entry.IsDirectory) { AddDirectory(entry.Path, entry.LastWriteTime); return; }
        ThrowIfUnavailable();
        using Stream content = entry.OpenContent();
        AddFile(entry.Path, content, entry.LastWriteTime, null, cancellationToken);
    }

    /// <summary>Writes the central directory and end records. Further additions are rejected. Dispose also finishes.</summary>
    public void Finish() {
        if (_finished) return;
        ThrowIfUnavailable();
        _finished = true;
        WriteEnd();
    }

    private void WriteEnd() {
        long start = _sink.Position;
        foreach (CentralRecord record in _records) WriteCentralHeader(record);
        long size = _sink.Position - start;
        long count = _records.Count;
        bool zip64 = count >= 0xFFFF || start >= Max32 || size >= Max32;
        var writer = new LittleEndianWriter();
        if (zip64) {
            long zip64Offset = _sink.Position;
            writer.UInt32(0x06064b50).UInt64(44).UInt16(MadeBy).UInt16(45).UInt32(0).UInt32(0)
                .UInt64((ulong)count).UInt64((ulong)count).UInt64((ulong)size).UInt64((ulong)start);
            writer.UInt32(0x07064b50).UInt32(0).UInt64((ulong)zip64Offset).UInt32(1);
        }
        writer.UInt32(0x06054b50).UInt16(0).UInt16(0)
            .UInt16(count >= 0xFFFF ? (ushort)0xFFFF : (ushort)count).UInt16(count >= 0xFFFF ? (ushort)0xFFFF : (ushort)count)
            .UInt32(size >= Max32 ? Max32 : (uint)size).UInt32(start >= Max32 ? Max32 : (uint)start).UInt16(0);
        writer.WriteTo(_sink);
        _sink.Flush();
    }

    /// <summary>Completes any open member, finishes the archive and releases the destination unless it was left open.</summary>
    public void Dispose() {
        if (_disposed) return;
        try {
            if (!_faulted) {
                _open?.Dispose();
                Finish();
            }
        } finally {
            _disposed = true;
            if (!_leaveOpen) _destination.Dispose();
        }
    }

    /// <summary>Writes the entries, in the given order, to a new ZIP stream.</summary>
    public static void Create(Stream destination, IEnumerable<ArchiveWriteEntry> entries, ZipArchiveWriterOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(entries);
        using var writer = new ZipArchiveWriter(destination, options, leaveOpen: true);
        foreach (ArchiveWriteEntry entry in entries) {
            cancellationToken.ThrowIfCancellationRequested();
            writer.Add(entry, cancellationToken);
        }
        writer.Finish();
    }

    /// <summary>Writes the entries, in the given order, to a new ZIP file. An existing file is not overwritten.</summary>
    public static void Create(string destinationPath, IEnumerable<ArchiveWriteEntry> entries, ZipArchiveWriterOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(entries);
        ArchiveWriteSource.WriteToFile(destinationPath, output => Create(output, entries, options, cancellationToken));
    }

    /// <summary>Archives a folder's files and subfolders in ordinal path order. Symbolic links and reparse points are rejected.</summary>
    public static void CreateFromDirectory(string sourceDirectory, Stream destination, ZipArchiveWriterOptions? options = null,
        CancellationToken cancellationToken = default) {
        options ??= new ZipArchiveWriterOptions();
        Create(destination, ArchiveWriteSource.FromDirectory(sourceDirectory, options, cancellationToken), options, cancellationToken);
    }

    /// <summary>Archives a folder into a new ZIP file. An existing file is not overwritten.</summary>
    public static void CreateFromDirectory(string sourceDirectory, string destinationPath, ZipArchiveWriterOptions? options = null,
        CancellationToken cancellationToken = default) {
        options ??= new ZipArchiveWriterOptions();
        List<ArchiveWriteEntry> entries = ArchiveWriteSource.FromDirectory(sourceDirectory, options, cancellationToken);
        ArchiveWriteSource.WriteToFile(destinationPath, output => Create(output, entries, options, cancellationToken));
    }

    private EntryStream Begin(string path, DateTimeOffset? lastWriteTime, ZipCompressionMethod? compression, long? knownLength) {
        ThrowIfUnavailable();
        ArgumentNullException.ThrowIfNull(path);
        ZipCompressionMethod method = compression ?? _options.Compression;
        if (method != ZipCompressionMethod.Stored && method != ZipCompressionMethod.Deflate)
            throw new ArgumentOutOfRangeException(nameof(compression), "Unsupported ZIP compression method.");
        if (path.EndsWith("/", StringComparison.Ordinal) || path.EndsWith("\\", StringComparison.Ordinal))
            throw new ArgumentException("File member names cannot end with a separator; use AddDirectory.", nameof(path));
        string normalized = _paths.Add(path);
        byte[] name = Encoding.UTF8.GetBytes(normalized);
        if (name.Length > 0xFFFF) throw new ArgumentException("The member name exceeds 65,535 UTF-8 bytes.", nameof(path));
        if (knownLength == 0) method = ZipCompressionMethod.Stored;
        var record = new CentralRecord(name, Flags(name, _options.UseDataDescriptors), (ushort)method, Dos(lastWriteTime), false, _sink.Position);
        Stream target;
        if (_options.UseDataDescriptors) {
            record.Zip64 = knownLength >= 0xFF000000L;
            WriteLocalHeader(record);
            target = _sink;
        } else target = new ArchiveSpillStream();
        _open = new EntryStream(this, record, target, method == ZipCompressionMethod.Deflate);
        return _open;
    }

    /// <summary>Abandons a member whose content could not be read. The archive stays without a central directory,
    /// so readers reject it instead of seeing a truncated member as complete.</summary>
    private void Abort(EntryStream entry) {
        _faulted = true; _open = null;
        entry.Abandon();
        if (entry.Target is ArchiveSpillStream spill) spill.Dispose();
    }

    private void Complete(EntryStream entry) {
        _open = null;
        CentralRecord record = entry.Record;
        if (_options.UseDataDescriptors) {
            long compressed = _sink.Position - entry.DataStart;
            record.Complete(entry.Crc32, compressed, entry.Uncompressed);
            if (!record.Zip64 && (compressed >= Max32 || entry.Uncompressed >= Max32))
                throw new IOException("A streamed ZIP member of unknown length exceeded 4 GiB; write it without data descriptors.");
            var writer = new LittleEndianWriter().UInt32(0x08074b50).UInt32(entry.Crc32);
            if (record.Zip64) writer.UInt64((ulong)compressed).UInt64((ulong)entry.Uncompressed);
            else writer.UInt32((uint)compressed).UInt32((uint)entry.Uncompressed);
            writer.WriteTo(_sink);
        } else {
            using var spill = (ArchiveSpillStream)entry.Target;
            if (entry.Uncompressed == 0) record.Method = 0;
            long compressed = record.Method == 0 && entry.Uncompressed == 0 ? 0 : spill.Length;
            record.Complete(entry.Crc32, compressed, entry.Uncompressed);
            record.Zip64 = compressed >= Max32 || entry.Uncompressed >= Max32;
            WriteLocalHeader(record);
            if (compressed > 0) spill.CopyTo(_sink, CancellationToken.None);
        }
        _records.Add(record);
    }

    private void WriteLocalHeader(CentralRecord record) {
        bool descriptor = (record.Flags & 0x0008) != 0;
        var writer = new LittleEndianWriter().UInt32(0x04034b50).UInt16(VersionNeeded(record)).UInt16(record.Flags)
            .UInt16(record.Method).UInt16(record.Time).UInt16(record.Date);
        if (descriptor) writer.UInt32(0).UInt32(record.Zip64 ? Max32 : 0).UInt32(record.Zip64 ? Max32 : 0);
        else writer.UInt32(record.Crc32).UInt32(record.Zip64 ? Max32 : (uint)record.Compressed).UInt32(record.Zip64 ? Max32 : (uint)record.Uncompressed);
        writer.UInt16((ushort)record.Name.Length).UInt16(record.Zip64 ? (ushort)20 : (ushort)0).Bytes(record.Name);
        if (record.Zip64) writer.UInt16(1).UInt16(16).UInt64(descriptor ? 0 : (ulong)record.Uncompressed).UInt64(descriptor ? 0 : (ulong)record.Compressed);
        writer.WriteTo(_sink);
        record.DataStart = _sink.Position;
    }

    private void WriteCentralHeader(CentralRecord record) {
        bool offset64 = record.Offset >= Max32;
        int extra = (record.Zip64 ? 16 : 0) + (offset64 ? 8 : 0);
        var writer = new LittleEndianWriter().UInt32(0x02014b50).UInt16(MadeBy).UInt16(VersionNeeded(record)).UInt16(record.Flags)
            .UInt16(record.Method).UInt16(record.Time).UInt16(record.Date).UInt32(record.Crc32)
            .UInt32(record.Zip64 ? Max32 : (uint)record.Compressed).UInt32(record.Zip64 ? Max32 : (uint)record.Uncompressed)
            .UInt16((ushort)record.Name.Length).UInt16(extra == 0 ? (ushort)0 : (ushort)(extra + 4)).UInt16(0).UInt16(0).UInt16(0)
            .UInt32(record.IsDirectory ? DirectoryAttributes : FileAttributes).UInt32(offset64 ? Max32 : (uint)record.Offset)
            .Bytes(record.Name);
        if (extra != 0) {
            writer.UInt16(1).UInt16((ushort)extra);
            if (record.Zip64) writer.UInt64((ulong)record.Uncompressed).UInt64((ulong)record.Compressed);
            if (offset64) writer.UInt64((ulong)record.Offset);
        }
        writer.WriteTo(_sink);
    }

    private static ushort VersionNeeded(CentralRecord record) =>
        record.Zip64 || record.Offset >= Max32 ? (ushort)45 : record.Method == 8 || record.IsDirectory ? (ushort)20 : (ushort)10;

    private static ushort Flags(byte[] name, bool descriptor) {
        ushort flags = descriptor ? (ushort)0x0008 : (ushort)0;
        foreach (byte value in name) if (value > 127) return (ushort)(flags | 0x0800);
        return flags;
    }

    private (ushort Time, ushort Date) Dos(DateTimeOffset? lastWriteTime) {
        DateTime value = (_options.Resolve(lastWriteTime) ?? DosEpoch).UtcDateTime;
        if (value < DosEpoch.UtcDateTime) value = DosEpoch.UtcDateTime;
        var latest = new DateTime(2107, 12, 31, 23, 59, 58, DateTimeKind.Utc);
        if (value > latest) value = latest;
        return ((ushort)(value.Hour << 11 | value.Minute << 5 | value.Second / 2),
            (ushort)((value.Year - 1980) << 9 | value.Month << 5 | value.Day));
    }

    private void ThrowIfUnavailable() {
        if (_disposed) throw new ObjectDisposedException(nameof(ZipArchiveWriter));
        if (_faulted) throw new InvalidOperationException("A previous member failed; the ZIP archive is incomplete.");
        if (_finished) throw new InvalidOperationException("The ZIP archive is already finished.");
        if (_open != null) throw new InvalidOperationException("Dispose the open member stream before adding another member.");
    }

    private sealed class CentralRecord {
        internal CentralRecord(byte[] name, ushort flags, ushort method, (ushort Time, ushort Date) dos, bool directory, long offset) {
            Name = name; Flags = flags; Method = method; Time = dos.Time; Date = dos.Date; IsDirectory = directory; Offset = offset;
        }
        internal byte[] Name { get; }
        internal ushort Flags { get; }
        internal ushort Method { get; set; }
        internal ushort Time { get; }
        internal ushort Date { get; }
        internal bool IsDirectory { get; }
        internal long Offset { get; }
        internal long DataStart { get; set; }
        internal bool Zip64 { get; set; }
        internal uint Crc32 { get; private set; }
        internal long Compressed { get; private set; }
        internal long Uncompressed { get; private set; }
        internal void Complete(uint crc, long compressed, long uncompressed) { Crc32 = crc; Compressed = compressed; Uncompressed = uncompressed; }
    }

    private sealed class EntryStream : Stream {
        private readonly ZipArchiveWriter _owner;
        private readonly ArchiveDeflateEncoder? _encoder;
        private uint _crc = OfficePngCrc32.Begin();
        private bool _closed;
        internal EntryStream(ZipArchiveWriter owner, CentralRecord record, Stream target, bool deflate) {
            _owner = owner; Record = record; Target = target; DataStart = owner._sink.Position;
            if (deflate) _encoder = new ArchiveDeflateEncoder(target);
        }
        internal CentralRecord Record { get; }
        internal Stream Target { get; }
        internal long DataStart { get; }
        internal long Uncompressed { get; private set; }
        internal uint Crc32 => OfficePngCrc32.Complete(_crc);
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => !_closed;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => Uncompressed; set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) {
            if (_closed) throw new ObjectDisposedException(nameof(EntryStream));
            ArgumentNullException.ThrowIfNull(buffer);
            if (offset < 0 || count < 0 || offset > buffer.Length - count) throw new ArgumentOutOfRangeException(nameof(count));
            if (count == 0) return;
            _crc = OfficePngCrc32.Append(_crc, buffer, offset, count);
            Uncompressed += count;
            if (_encoder != null) _encoder.Write(buffer, offset, count);
            else Target.Write(buffer, offset, count);
        }
        internal void Abandon() => _closed = true;
        protected override void Dispose(bool disposing) {
            if (disposing && !_closed) {
                _closed = true;
                bool emptyStaged = Uncompressed == 0 && Target is ArchiveSpillStream;
                if (_encoder != null && !emptyStaged) _encoder.Finish();
                _owner.Complete(this);
            }
            base.Dispose(disposing);
        }
    }

    private sealed class CountingStream : Stream {
        private readonly Stream _inner;
        internal CountingStream(Stream inner) => _inner = inner;
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => throw new NotSupportedException();
        public override long Position { get; set; }
        public override void Flush() => _inner.Flush();
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) { _inner.Write(buffer, offset, count); Position += count; }
    }
}

/// <summary>Little-endian record builder for archive headers.</summary>
internal sealed class LittleEndianWriter {
    private readonly MemoryStream _buffer = new();
    internal LittleEndianWriter UInt16(ushort value) { _buffer.WriteByte((byte)value); _buffer.WriteByte((byte)(value >> 8)); return this; }
    internal LittleEndianWriter UInt32(uint value) { UInt16((ushort)value); return UInt16((ushort)(value >> 16)); }
    internal LittleEndianWriter UInt64(ulong value) { UInt32((uint)value); return UInt32((uint)(value >> 32)); }
    internal LittleEndianWriter Bytes(byte[] value) { _buffer.Write(value, 0, value.Length); return this; }
    internal void WriteTo(Stream destination) => _buffer.WriteTo(destination);
}
