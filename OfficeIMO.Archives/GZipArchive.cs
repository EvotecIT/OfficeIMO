using System.IO.Compression;
using System.Text;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;

namespace OfficeIMO.Archives;

/// <summary>Settings for writing a gzip (RFC 1952) member.</summary>
public sealed class GZipWriterOptions {
    /// <summary>Optional original file name (FNAME), stored as ISO-8859-1. Omitted by default.</summary>
    public string? FileName { get; set; }
    /// <summary>Optional modification time (MTIME). Defaults to zero, which RFC 1952 defines as "no time stamp",
    /// so the output depends only on the content.</summary>
    public DateTimeOffset? ModificationTime { get; set; }
}

/// <summary>The outcome of decompressing a gzip stream.</summary>
public sealed class GZipReadResult {
    internal GZipReadResult(string? fileName, DateTimeOffset? modificationTime, int members, long length) {
        FileName = fileName; ModificationTime = modificationTime; MemberCount = members; Length = length;
    }
    /// <summary>The first member's FNAME field, or null when absent.</summary>
    public string? FileName { get; }
    /// <summary>The first member's MTIME field, or null when it is zero.</summary>
    public DateTimeOffset? ModificationTime { get; }
    /// <summary>Number of concatenated gzip members read.</summary>
    public int MemberCount { get; }
    /// <summary>Total decompressed bytes written to the destination.</summary>
    public long Length { get; }
}

/// <summary>Reads and writes gzip (RFC 1952) streams. Writing uses the same deterministic Deflate encoder as the ZIP writer;
/// reading verifies every member's CRC-32 and length and accepts concatenated members.</summary>
public static class GZipArchive {
    /// <summary>Returns a stream that compresses everything written to it as one gzip member. Disposing it writes the trailer.</summary>
    public static Stream CreateWriter(Stream destination, GZipWriterOptions? options = null, bool leaveOpen = false) {
        ArgumentNullException.ThrowIfNull(destination);
        if (!destination.CanWrite) throw new ArgumentException("A writable destination stream is required.", nameof(destination));
        return new GZipWriteStream(destination, options ?? new GZipWriterOptions(), leaveOpen);
    }

    /// <summary>Compresses the source from its current position to its end as one gzip member. Both streams stay open.</summary>
    public static void Compress(Stream source, Stream destination, GZipWriterOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(source);
        using Stream writer = CreateWriter(destination, options, leaveOpen: true);
        var buffer = new byte[81920];
        int read;
        while ((read = source.Read(buffer, 0, buffer.Length)) > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            writer.Write(buffer, 0, read);
        }
    }

    /// <summary>Decompresses every member of a gzip stream into the destination, verifying checksums and lengths.
    /// Input and output are bounded by <see cref="ArchiveReadOptions.MaxInputBytes"/> and <see cref="ArchiveReadOptions.MaxEntryBytes"/>.
    /// Nonseekable input is snapshotted first. The source stays open.</summary>
    /// <exception cref="InvalidDataException">The input is not gzip, is truncated, or fails a checksum.</exception>
    /// <exception cref="OfficeInspectionResourceLimitException">A configured limit is exceeded.</exception>
    public static GZipReadResult Decompress(Stream source, Stream destination, ArchiveReadOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(destination);
        var limits = (options ?? new ArchiveReadOptions()).Limits();
        using Stream? snapshot = source.CanSeek ? null : ArchiveReader.Snapshot(source, limits, cancellationToken);
        Stream input = snapshot ?? source;
        if (snapshot == null) {
            ArchiveLimits.Check(input.Length - input.Position, limits.Input, nameof(ArchiveReadOptions.MaxInputBytes));
            limits.Operation.AdmitInput(input.Length - input.Position);
        }
        return GZipContent.Decompress(input, destination, limits.Entry, limits, nameof(ArchiveReadOptions.MaxEntryBytes), cancellationToken);
    }
}

internal static class GZipContent {
    internal const int HeaderTextLimit = 64 * 1024;

    internal static bool HasSignature(byte[] bytes, int length) => length >= 3 && bytes[0] == 0x1f && bytes[1] == 0x8b && bytes[2] == 8;

    internal sealed class Header {
        internal string? FileName;
        internal DateTimeOffset? ModificationTime;
    }

    /// <summary>Reads one member header at the current position, leaving the stream at the first Deflate byte.</summary>
    internal static Header ReadHeader(Stream source) {
        var bytes = new List<byte>(16);
        int Next() {
            int value = source.ReadByte();
            if (value < 0) throw new InvalidDataException("The gzip header is truncated.");
            bytes.Add((byte)value);
            return value;
        }
        if (Next() != 0x1f || Next() != 0x8b) throw new InvalidDataException("The gzip signature is missing.");
        if (Next() != 8) throw new NotSupportedException("The gzip member uses a compression method other than Deflate.");
        int flags = Next();
        if ((flags & 0xe0) != 0) throw new InvalidDataException("The gzip header sets reserved flags.");
        uint mtime = (uint)(Next() | Next() << 8 | Next() << 16 | Next() << 24);
        Next(); Next();
        var header = new Header { ModificationTime = mtime == 0 ? null : DateTimeOffset.FromUnixTimeSeconds(mtime) };
        if ((flags & 0x04) != 0) {
            int extra = Next() | Next() << 8;
            for (int index = 0; index < extra; index++) Next();
        }
        if ((flags & 0x08) != 0) header.FileName = ReadText(Next, "file name");
        if ((flags & 0x10) != 0) ReadText(Next, "comment");
        if ((flags & 0x02) != 0) {
            uint expected = OfficePngCrc32.Complete(OfficePngCrc32.Append(OfficePngCrc32.Begin(), bytes.ToArray(), 0, bytes.Count)) & 0xffff;
            int stored = Next() | Next() << 8;
            if (stored != expected) throw new InvalidDataException("The gzip header checksum does not match.");
        }
        return header;
    }

    private static string ReadText(Func<int> next, string field) {
        var text = new StringBuilder();
        int value;
        while ((value = next()) != 0) {
            if (text.Length == HeaderTextLimit) throw new InvalidDataException("The gzip " + field + " is too long.");
            text.Append((char)value);
        }
        return text.ToString();
    }

    /// <summary>Decompresses all members from the current position of a seekable stream.</summary>
    internal static GZipReadResult Decompress(Stream source, Stream destination, long maximumOutput, ArchiveLimits limits,
        string limitName, CancellationToken token) {
        long length = source.Length;
        long total = 0;
        int members = 0;
        Header? first = null;
        var buffer = new byte[81920];
        while (true) {
            token.ThrowIfCancellationRequested();
            Header header = ReadHeader(source);
            first ??= header;
            members++;
            long start = source.Position;
            long remaining = maximumOutput - total;
            if (!OfficeDeflateStreamValidator.TryMeasure(source, length - start, remaining, out long consumed, out long expected,
                    out bool exceeded, token)) {
                if (exceeded) throw ArchiveLimits.Exceeded(limitName, maximumOutput);
                throw new InvalidDataException("The gzip member contains truncated or invalid Deflate data.");
            }
            limits.Operation.Decode(expected, consumed + expected);
            uint crc = OfficePngCrc32.Begin();
            long produced = 0;
            using (var window = new SeekableReadWindowStream(source, start, consumed))
            using (var inflater = new DeflateStream(window, CompressionMode.Decompress, leaveOpen: true)) {
                int read;
                while ((read = inflater.Read(buffer, 0, buffer.Length)) > 0) {
                    token.ThrowIfCancellationRequested();
                    if (read > expected - produced) throw new InvalidDataException("The gzip member expands beyond its validated length.");
                    crc = OfficePngCrc32.Append(crc, buffer, 0, read);
                    destination.Write(buffer, 0, read);
                    produced += read;
                }
            }
            if (produced != expected) throw new InvalidDataException("The gzip member is shorter than its validated length.");
            source.Position = start + consumed;
            var trailer = new byte[8];
            if (ReadFully(source, trailer) != 8) throw new InvalidDataException("The gzip trailer is truncated.");
            uint storedCrc = (uint)(trailer[0] | trailer[1] << 8 | trailer[2] << 16 | trailer[3] << 24);
            uint storedSize = (uint)(trailer[4] | trailer[5] << 8 | trailer[6] << 16 | trailer[7] << 24);
            if (storedCrc != OfficePngCrc32.Complete(crc)) throw new InvalidDataException("The gzip member checksum does not match.");
            if (storedSize != unchecked((uint)produced)) throw new InvalidDataException("The gzip member length does not match its trailer.");
            total += produced;
            if (source.Position == length) break;
            int b1 = source.ReadByte(), b2 = source.ReadByte();
            source.Position -= b2 < 0 ? 1 : 2;
            if (b1 == 0x1f && b2 == 0x8b) continue;
            // Tape and block devices pad with zeros after the last member; anything else is not gzip data.
            while (true) {
                int read = source.Read(buffer, 0, buffer.Length);
                if (read == 0) break;
                for (int index = 0; index < read; index++)
                    if (buffer[index] != 0) throw new InvalidDataException("Unexpected data follows the last gzip member.");
            }
            break;
        }
        return new GZipReadResult(first!.FileName, first.ModificationTime, members, total);
    }

    /// <summary>Decompresses at most <paramref name="count"/> bytes of the first member for format detection.</summary>
    internal static int Peek(Stream source, byte[] destination, int count) {
        try {
            ReadHeader(source);
            using var inflater = new DeflateStream(source, CompressionMode.Decompress, leaveOpen: true);
            int total = 0, read;
            while (total < count && (read = inflater.Read(destination, total, count - total)) > 0) total += read;
            return total;
        } catch (Exception error) when (error is InvalidDataException or NotSupportedException or EndOfStreamException) {
            return 0;
        }
    }

    private static int ReadFully(Stream source, byte[] buffer) {
        int total = 0, read;
        while (total < buffer.Length && (read = source.Read(buffer, total, buffer.Length - total)) > 0) total += read;
        return total;
    }
}

internal sealed class GZipWriteStream : Stream {
    private readonly Stream _destination;
    private readonly bool _leaveOpen;
    private readonly ArchiveDeflateEncoder _encoder;
    private uint _crc = OfficePngCrc32.Begin();
    private long _length;
    private bool _closed;

    internal GZipWriteStream(Stream destination, GZipWriterOptions options, bool leaveOpen) {
        _destination = destination; _leaveOpen = leaveOpen;
        byte[]? name = null;
        if (options.FileName != null) {
            if (options.FileName.Length == 0 || options.FileName.Any(character => character == 0 || character > 0xff))
                throw new ArgumentException("The gzip file name must be nonempty ISO-8859-1 text without NUL characters.", nameof(options));
            name = options.FileName.Select(character => (byte)character).ToArray();
        }
        long seconds = options.ModificationTime?.ToUnixTimeSeconds() ?? 0;
        if (seconds < 0 || seconds > uint.MaxValue) throw new ArgumentOutOfRangeException(nameof(options), "The gzip time must fall between 1970 and 2106.");
        var header = new LittleEndianWriter();
        header.Bytes(new byte[] { 0x1f, 0x8b, 8, name == null ? (byte)0 : (byte)0x08 }).UInt32((uint)seconds).Bytes(new byte[] { 0, 255 });
        if (name != null) header.Bytes(name).Bytes(new byte[] { 0 });
        header.WriteTo(destination);
        _encoder = new ArchiveDeflateEncoder(destination);
    }

    public override bool CanRead => false;
    public override bool CanSeek => false;
    public override bool CanWrite => !_closed;
    public override long Length => throw new NotSupportedException();
    public override long Position { get => _length; set => throw new NotSupportedException(); }
    public override void Flush() { }
    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) {
        if (_closed) throw new ObjectDisposedException(nameof(GZipWriteStream));
        ArgumentNullException.ThrowIfNull(buffer);
        if (offset < 0 || count < 0 || offset > buffer.Length - count) throw new ArgumentOutOfRangeException(nameof(count));
        if (count == 0) return;
        _crc = OfficePngCrc32.Append(_crc, buffer, offset, count);
        _length += count;
        _encoder.Write(buffer, offset, count);
    }
    protected override void Dispose(bool disposing) {
        if (disposing && !_closed) {
            _closed = true;
            try {
                _encoder.Finish();
                new LittleEndianWriter().UInt32(OfficePngCrc32.Complete(_crc)).UInt32(unchecked((uint)_length)).WriteTo(_destination);
                _destination.Flush();
            } finally {
                if (!_leaveOpen) _destination.Dispose();
            }
        }
        base.Dispose(disposing);
    }
}
