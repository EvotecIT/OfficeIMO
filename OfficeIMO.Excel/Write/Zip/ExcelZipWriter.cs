#if NET8_0_OR_GREATER
using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel;

/// <summary>A forward-writing ZIP archive (PKWARE APPNOTE 6.3, section 4.3) for newly created packages.</summary>
/// <remarks>
/// <para>Each entry's local header is written when its first bytes arrive, so a destination that fails before any
/// content leaves the caller's own error in charge. On a seekable destination the header's CRC and sizes are filled
/// in afterwards; on a forward-only one the entry carries a data descriptor (general purpose bit 3).</para>
/// <para>ZIP64 extra fields and end records (section 4.5.3 and 4.3.14) are added only for values that do not fit
/// in 32 bits. An entry whose sizes outgrow its local header on a seekable destination switches to a ZIP64 data
/// descriptor, which the central directory then describes.</para>
/// </remarks>
internal sealed class ExcelZipWriter : ExcelPackageArchive {
    private const ushort DescriptorFlag = 1 << 3;
    private const ushort Utf8NameFlag = 1 << 11;
    private const ushort Deflated = 8;
    private const ushort Stored = 0;
    private const ushort Version20 = 20;
    private const ushort Version45 = 45;

    private readonly Stream _destination;
    private readonly bool _seekable;
    private readonly long _origin;
    private readonly ushort _dosTime;
    private readonly ushort _dosDate;
    private readonly CountingSink _sink;
    private readonly List<Part> _parts = new();
    private PartStream? _current;
    private long _offset;
    private bool _faulted;
    private bool _closed;

    internal ExcelZipWriter(Stream destination, DateTime modified) {
        _destination = destination;
        _seekable = destination.CanSeek;
        if (_seekable) _origin = destination.Position;
        DateTime time = modified.Year < 1980 ? new DateTime(1980, 1, 1) : modified.Year > 2107 ? new DateTime(2107, 12, 31, 23, 59, 58) : modified;
        _dosTime = (ushort)((time.Hour << 11) | (time.Minute << 5) | (time.Second / 2));
        _dosDate = (ushort)(((time.Year - 1980) << 9) | (time.Month << 5) | time.Day);
        _sink = new CountingSink(this);
    }

    internal override Stream OpenEntry(string name, CompressionLevel level) {
        if (_closed) throw new ObjectDisposedException(nameof(ExcelZipWriter));
        if (_current != null) throw new InvalidOperationException("Close the previous package entry before opening another.");
        if (_faulted) throw new InvalidOperationException("A package entry failed; the archive cannot take more entries.");
        var part = new Part(name, level);
        if (!_seekable) part.Flags |= DescriptorFlag;
        return _current = new PartStream(this, part);
    }

    /// <summary>
    /// Deflates the bytes <paramref name="write"/> produces, as a <see cref="CompressionLevel.Fastest"/> entry would be,
    /// without touching an archive, so several parts can be prepared at once and appended in order.
    /// </summary>
    internal static PreparedEntry Prepare(Action<Stream> write) {
        var buffer = new MemoryStream();
        using (var stream = new PreparingStream(buffer)) {
            write(stream);
            stream.Complete();
            return new PreparedEntry(buffer, stream.Crc, stream.Size);
        }
    }

    /// <summary>Appends a part that <see cref="Prepare"/> compressed; the entry is the one <see cref="OpenEntry"/> would write.</summary>
    internal void WritePrepared(string name, PreparedEntry entry) {
        if (_closed) throw new ObjectDisposedException(nameof(ExcelZipWriter));
        if (_current != null) throw new InvalidOperationException("Close the previous package entry before opening another.");
        if (_faulted) throw new InvalidOperationException("A package entry failed; the archive cannot take more entries.");
        var part = new Part(name, CompressionLevel.Fastest);
        if (!_seekable) part.Flags |= DescriptorFlag;
        try {
            WriteLocalHeader(part);
            long contentStart = _offset;
            _sink.Write(entry.Content.GetBuffer(), 0, checked((int)entry.Content.Length));
            part.Crc = entry.Crc;
            part.Size = entry.Size;
            part.CompressedSize = _offset - contentStart;
        } catch {
            FailPart();
            throw;
        }
        CompletePart(part);
    }

    /// <summary>A part deflated away from the archive, with the CRC-32 and size of its uncompressed bytes.</summary>
    internal sealed class PreparedEntry {
        internal PreparedEntry(MemoryStream content, uint crc, long size) {
            Content = content;
            Crc = crc;
            Size = size;
        }

        internal MemoryStream Content { get; }
        internal uint Crc { get; }
        internal long Size { get; }
    }

    private sealed class PreparingStream(MemoryStream output) : Stream {
        private readonly ExcelDeflateEncoder _encoder = new(output);
        private uint _crc = OfficePngCrc32.Begin();
        private bool _completed;

        internal uint Crc => OfficePngCrc32.Complete(_crc);
        internal long Size { get; private set; }

        public override void Write(byte[] buffer, int offset, int count) {
            ValidateBufferArguments(buffer, offset, count);
            Write(new ReadOnlySpan<byte>(buffer, offset, count));
        }

        public override void Write(ReadOnlySpan<byte> buffer) {
            if (_completed) throw new ObjectDisposedException(nameof(PreparingStream));
            _crc = OfficePngCrc32.Append(_crc, buffer);
            Size += buffer.Length;
            _encoder.Write(buffer);
        }

        internal void Complete() {
            if (_completed) return;
            _completed = true;
            _encoder.Finish();
        }

        protected override void Dispose(bool disposing) {
            if (disposing) _encoder.Dispose();
            base.Dispose(disposing);
        }

        public override void Flush() { }
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => !_completed;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
    }

    /// <summary>Writes the central directory once every entry has completed. A faulted archive stays unfinished.</summary>
    public override void Dispose() {
        if (_closed) return;
        _closed = true;
        AbandonCompression();
        if (_current != null || _faulted) return;
        long directoryStart = _offset;
        using var directory = new MemoryStream(64 + _parts.Count * 80);
        foreach (Part part in _parts) WriteCentralHeader(directory, part);
        WriteEnd(directory, directoryStart, directory.Length);
        _sink.Write(directory.GetBuffer(), 0, checked((int)directory.Length));
    }

    private void WriteLocalHeader(Part part) {
        part.HeaderOffset = _offset;
        byte[] header = new byte[30 + part.Name.Length];
        Span<byte> span = header;
        BinaryPrimitives.WriteUInt32LittleEndian(span, 0x0403_4B50);
        BinaryPrimitives.WriteUInt16LittleEndian(span.Slice(4), Version20);
        BinaryPrimitives.WriteUInt16LittleEndian(span.Slice(6), part.Flags);
        BinaryPrimitives.WriteUInt16LittleEndian(span.Slice(8), part.Method);
        BinaryPrimitives.WriteUInt16LittleEndian(span.Slice(10), _dosTime);
        BinaryPrimitives.WriteUInt16LittleEndian(span.Slice(12), _dosDate);
        // CRC-32 and both sizes (offsets 14-25) stay zero until the entry completes.
        BinaryPrimitives.WriteUInt16LittleEndian(span.Slice(26), checked((ushort)part.Name.Length));
        part.Name.CopyTo(header, 30);
        _sink.Write(header, 0, header.Length);
    }

    private void CompletePart(Part part) {
        _current = null;
        bool large = part.NeedsZip64Sizes;
        if ((part.Flags & DescriptorFlag) == 0 && large) {
            part.Flags |= DescriptorFlag;
            byte[] versionAndFlags = new byte[4];
            BinaryPrimitives.WriteUInt16LittleEndian(versionAndFlags, Version45);
            BinaryPrimitives.WriteUInt16LittleEndian(versionAndFlags.AsSpan(2), part.Flags);
            Overwrite(part.HeaderOffset + 4, versionAndFlags);
        }
        if ((part.Flags & DescriptorFlag) != 0) {
            byte[] descriptor = new byte[large ? 24 : 16];
            BinaryPrimitives.WriteUInt32LittleEndian(descriptor, 0x0807_4B50);
            BinaryPrimitives.WriteUInt32LittleEndian(descriptor.AsSpan(4), part.Crc);
            if (large) {
                BinaryPrimitives.WriteUInt64LittleEndian(descriptor.AsSpan(8), (ulong)part.CompressedSize);
                BinaryPrimitives.WriteUInt64LittleEndian(descriptor.AsSpan(16), (ulong)part.Size);
            } else {
                BinaryPrimitives.WriteUInt32LittleEndian(descriptor.AsSpan(8), (uint)part.CompressedSize);
                BinaryPrimitives.WriteUInt32LittleEndian(descriptor.AsSpan(12), (uint)part.Size);
            }
            _sink.Write(descriptor, 0, descriptor.Length);
        } else {
            byte[] crcAndSizes = new byte[12];
            BinaryPrimitives.WriteUInt32LittleEndian(crcAndSizes, part.Crc);
            BinaryPrimitives.WriteUInt32LittleEndian(crcAndSizes.AsSpan(4), (uint)part.CompressedSize);
            BinaryPrimitives.WriteUInt32LittleEndian(crcAndSizes.AsSpan(8), (uint)part.Size);
            Overwrite(part.HeaderOffset + 14, crcAndSizes);
        }
        _parts.Add(part);
    }

    private void FailPart() {
        _current = null;
        _faulted = true;
    }

    private void Overwrite(long offset, byte[] bytes) {
        _destination.Seek(_origin + offset, SeekOrigin.Begin);
        _destination.Write(bytes, 0, bytes.Length);
        _destination.Seek(_origin + _offset, SeekOrigin.Begin);
    }

    private void WriteCentralHeader(MemoryStream directory, Part part) {
        bool largeSizes = part.NeedsZip64Sizes;
        bool largeOffset = part.HeaderOffset >= uint.MaxValue;
        int zip64Length = (largeSizes ? 16 : 0) + (largeOffset ? 8 : 0);
        int extraLength = zip64Length == 0 ? 0 : 4 + zip64Length;
        byte[] header = new byte[46 + part.Name.Length + extraLength];
        Span<byte> span = header;
        BinaryPrimitives.WriteUInt32LittleEndian(span, 0x0201_4B50);
        BinaryPrimitives.WriteUInt16LittleEndian(span.Slice(4), Version45);
        BinaryPrimitives.WriteUInt16LittleEndian(span.Slice(6), zip64Length == 0 ? Version20 : Version45);
        BinaryPrimitives.WriteUInt16LittleEndian(span.Slice(8), part.Flags);
        BinaryPrimitives.WriteUInt16LittleEndian(span.Slice(10), part.Method);
        BinaryPrimitives.WriteUInt16LittleEndian(span.Slice(12), _dosTime);
        BinaryPrimitives.WriteUInt16LittleEndian(span.Slice(14), _dosDate);
        BinaryPrimitives.WriteUInt32LittleEndian(span.Slice(16), part.Crc);
        BinaryPrimitives.WriteUInt32LittleEndian(span.Slice(20), largeSizes ? uint.MaxValue : (uint)part.CompressedSize);
        BinaryPrimitives.WriteUInt32LittleEndian(span.Slice(24), largeSizes ? uint.MaxValue : (uint)part.Size);
        BinaryPrimitives.WriteUInt16LittleEndian(span.Slice(28), checked((ushort)part.Name.Length));
        BinaryPrimitives.WriteUInt16LittleEndian(span.Slice(30), checked((ushort)extraLength));
        // Comment length, disk number, internal and external attributes (32-41) stay zero.
        BinaryPrimitives.WriteUInt32LittleEndian(span.Slice(42), largeOffset ? uint.MaxValue : (uint)part.HeaderOffset);
        part.Name.CopyTo(header, 46);
        if (zip64Length != 0) {
            Span<byte> extra = span.Slice(46 + part.Name.Length);
            BinaryPrimitives.WriteUInt16LittleEndian(extra, 0x0001);
            BinaryPrimitives.WriteUInt16LittleEndian(extra.Slice(2), (ushort)zip64Length);
            int field = 4;
            // Section 4.5.3: only the fields set to 0xFFFFFFFF appear, in this order.
            if (largeSizes) {
                BinaryPrimitives.WriteUInt64LittleEndian(extra.Slice(field), (ulong)part.Size);
                BinaryPrimitives.WriteUInt64LittleEndian(extra.Slice(field + 8), (ulong)part.CompressedSize);
                field += 16;
            }
            if (largeOffset) BinaryPrimitives.WriteUInt64LittleEndian(extra.Slice(field), (ulong)part.HeaderOffset);
        }
        directory.Write(header, 0, header.Length);
    }

    private void WriteEnd(MemoryStream directory, long directoryStart, long directoryLength) {
        int count = _parts.Count;
        bool zip64 = count >= ushort.MaxValue || directoryStart >= uint.MaxValue || directoryLength >= uint.MaxValue;
        if (zip64) {
            long recordOffset = directoryStart + directoryLength;
            byte[] record = new byte[56 + 20];
            Span<byte> span = record;
            BinaryPrimitives.WriteUInt32LittleEndian(span, 0x0606_4B50);
            BinaryPrimitives.WriteUInt64LittleEndian(span.Slice(4), 56 - 12);
            BinaryPrimitives.WriteUInt16LittleEndian(span.Slice(12), Version45);
            BinaryPrimitives.WriteUInt16LittleEndian(span.Slice(14), Version45);
            BinaryPrimitives.WriteUInt64LittleEndian(span.Slice(24), (ulong)count);
            BinaryPrimitives.WriteUInt64LittleEndian(span.Slice(32), (ulong)count);
            BinaryPrimitives.WriteUInt64LittleEndian(span.Slice(40), (ulong)directoryLength);
            BinaryPrimitives.WriteUInt64LittleEndian(span.Slice(48), (ulong)directoryStart);
            // Locator: disk with the ZIP64 end record, its offset, total disks.
            BinaryPrimitives.WriteUInt32LittleEndian(span.Slice(56), 0x0706_4B50);
            BinaryPrimitives.WriteUInt64LittleEndian(span.Slice(64), (ulong)recordOffset);
            BinaryPrimitives.WriteUInt32LittleEndian(span.Slice(72), 1);
            directory.Write(record, 0, record.Length);
        }
        byte[] end = new byte[22];
        ushort entries = (ushort)Math.Min(count, ushort.MaxValue);
        BinaryPrimitives.WriteUInt32LittleEndian(end, 0x0605_4B50);
        BinaryPrimitives.WriteUInt16LittleEndian(end.AsSpan(8), entries);
        BinaryPrimitives.WriteUInt16LittleEndian(end.AsSpan(10), entries);
        BinaryPrimitives.WriteUInt32LittleEndian(end.AsSpan(12), directoryLength >= uint.MaxValue ? uint.MaxValue : (uint)directoryLength);
        BinaryPrimitives.WriteUInt32LittleEndian(end.AsSpan(16), directoryStart >= uint.MaxValue ? uint.MaxValue : (uint)directoryStart);
        directory.Write(end, 0, end.Length);
    }

    private sealed class Part {
        internal Part(string name, CompressionLevel level) {
            Name = Encoding.UTF8.GetBytes(name);
            Level = level;
            Method = level == CompressionLevel.NoCompression ? Stored : Deflated;
            // Bits 1-2 record the deflate option; bit 11 marks a UTF-8 name.
            Flags = level switch { CompressionLevel.Fastest => 0x0006, CompressionLevel.SmallestSize => 0x0002, _ => 0 };
            if (Name.Length != name.Length) Flags |= Utf8NameFlag;
        }

        internal byte[] Name { get; }
        internal CompressionLevel Level { get; }
        internal ushort Method { get; }
        internal ushort Flags { get; set; }
        internal long HeaderOffset { get; set; } = -1;
        internal uint Crc { get; set; }
        internal long CompressedSize { get; set; }
        internal long Size { get; set; }
        internal bool NeedsZip64Sizes => Size >= uint.MaxValue || CompressedSize >= uint.MaxValue;
    }

    /// <summary>Forwards to the destination and advances the archive offset, so offsets stay exact on forward-only outputs.</summary>
    private sealed class CountingSink(ExcelZipWriter owner) : Stream {
        public override void Write(byte[] buffer, int offset, int count) {
            owner._destination.Write(buffer, offset, count);
            owner._offset += count;
        }
        public override void Write(ReadOnlySpan<byte> buffer) {
            owner._destination.Write(buffer);
            owner._offset += buffer.Length;
        }
        public override void Flush() => owner._destination.Flush();
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
    }

    /// <summary>Hashes and compresses one entry as it is written; dispose completes it.</summary>
    private sealed class PartStream(ExcelZipWriter owner, Part part) : Stream {
        private ExcelDeflateEncoder? _encoder;
        private DeflateStream? _deflate;
        private uint _crc = OfficePngCrc32.Begin();
        private long _size;
        private long _contentStart;
        private bool _failed;
        private bool _disposed;

        public override void Write(byte[] buffer, int offset, int count) {
            ValidateBufferArguments(buffer, offset, count);
            Write(new ReadOnlySpan<byte>(buffer, offset, count));
        }

        public override void Write(ReadOnlySpan<byte> buffer) {
            if (_disposed) throw new ObjectDisposedException(nameof(PartStream));
            try {
                Start();
                _crc = OfficePngCrc32.Append(_crc, buffer);
                _size += buffer.Length;
                if (_encoder != null) _encoder.Write(buffer);
                else if (_deflate != null) _deflate.Write(buffer);
                else owner._sink.Write(buffer);
            } catch {
                _failed = true;
                throw;
            }
        }

        private void Start() {
            if (part.HeaderOffset >= 0) return;
            owner.WriteLocalHeader(part);
            _contentStart = owner._offset;
            if (part.Method == Stored) return;
            if (part.Level == CompressionLevel.Fastest) _encoder = new ExcelDeflateEncoder(owner._sink);
            else _deflate = new DeflateStream(owner._sink, part.Level, leaveOpen: true);
        }

        protected override void Dispose(bool disposing) {
            if (_disposed) return;
            _disposed = true;
            try {
                if (!disposing) return;
                if (_failed) { owner.FailPart(); return; }
                try {
                    Start();
                    _encoder?.Finish();
                    _deflate?.Dispose();
                } catch {
                    owner.FailPart();
                    throw;
                }
                part.Crc = OfficePngCrc32.Complete(_crc);
                part.Size = _size;
                part.CompressedSize = owner._offset - _contentStart;
                owner.CompletePart(part);
            } finally {
                _encoder?.Dispose();
                base.Dispose(disposing);
            }
        }

        public override void Flush() { }
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => !_disposed;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
    }
}
#endif
