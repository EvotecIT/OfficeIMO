using System.IO.Compression;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Core.Internal.OpenXml;

/// <summary>
/// Read-only ZIP container for Open XML packages. The central directory is scanned once by
/// <see cref="OfficeArchiveSafety"/>, which rejects layouts other readers could resolve differently; members are
/// inflated on demand with the runtime's <see cref="DeflateStream"/> into exactly sized buffers and checked against
/// their declared length and CRC-32.
/// </summary>
internal sealed class OfficeZipPackageReader : IDisposable {
    private const int MaximumEntries = 100_000;

    private readonly Stream _source;
    private readonly bool _leaveOpen;
    private readonly byte[]? _buffer;
    private readonly int _bufferOrigin;
    private readonly OfficeArchiveSafety.ZipCentralDirectoryRecord[] _records;
    private readonly Dictionary<string, int> _entries;
    private readonly long _maximumEntryLength;

    private OfficeZipPackageReader(Stream source, bool leaveOpen, byte[]? buffer, int bufferOrigin,
        OfficeArchiveSafety.ZipCentralDirectoryRecord[] records, Dictionary<string, int> entries, long maximumEntryLength) {
        _source = source;
        _leaveOpen = leaveOpen;
        _buffer = buffer;
        _bufferOrigin = bufferOrigin;
        _records = records;
        _entries = entries;
        _maximumEntryLength = maximumEntryLength;
    }

    /// <summary>Number of members, including the content-types catalog and relationship parts.</summary>
    internal int Count => _records.Length;

    /// <summary>Opens the ZIP archive that starts at the current position of a readable, seekable stream.</summary>
    /// <exception cref="InvalidDataException">The archive is malformed, ambiguous or repeats a member name.</exception>
    internal static OfficeZipPackageReader Open(Stream source, bool leaveOpen = false, long maximumEntryLength = 1L << 30) {
        ArgumentNullException.ThrowIfNull(source);
        if (!source.CanRead || !source.CanSeek) throw new ArgumentException("The package stream must be readable and seekable.", nameof(source));
        long origin = source.Position;
        var records = new List<OfficeArchiveSafety.ZipCentralDirectoryRecord>();
        OfficeArchiveSafety.ZipCentralDirectoryScanResult scan = OfficeArchiveSafety.ScanZipCentralDirectoryForExtraction(
            source, source.Length - origin, MaximumEntries, records);
        if (!scan.IsValid) throw new InvalidDataException(scan.Error ?? "The package is not a valid ZIP archive.");
        if (scan.LimitExceeded) throw new InvalidDataException("The package has more members than an Open XML reader accepts.");

        var entries = new Dictionary<string, int>(records.Count, OfficeAsciiIgnoreCaseComparer.Instance);
        for (int index = 0; index < records.Count; index++) {
            string name = DecodeName(records[index].NameBytes);
            // OPC part names compare without regard to ASCII case, so two such members name the same part.
            if (entries.ContainsKey(name)) throw new InvalidDataException("The package repeats the member '" + name + "'.");
            entries.Add(name, index);
        }

        byte[]? buffer = null;
        int bufferOrigin = 0;
        if (source is MemoryStream memory && memory.TryGetBuffer(out ArraySegment<byte> segment) && segment.Array != null) {
            buffer = segment.Array;
            bufferOrigin = segment.Offset;
        }
        return new OfficeZipPackageReader(source, leaveOpen, buffer, bufferOrigin, records.ToArray(), entries, maximumEntryLength);
    }

    /// <summary>Member names in central-directory order.</summary>
    internal IEnumerable<string> Names => _entries.Keys;

    /// <summary>Returns whether the archive has a member with this name (ASCII case-insensitive).</summary>
    internal bool Contains(string name) => _entries.ContainsKey(name);

    /// <summary>The uncompressed length a member declares in the central directory.</summary>
    internal bool TryGetLength(string name, out long length) {
        bool found = _entries.TryGetValue(name, out int index);
        length = found ? _records[index].UncompressedSize : 0;
        return found;
    }

    /// <summary>Inflates one member into a buffer the caller disposes.</summary>
    /// <returns><see langword="false"/> when no member has this name.</returns>
    /// <exception cref="InvalidDataException">The member's local header, length or checksum disagrees with the directory.</exception>
    internal bool TryRead(string name, out OfficePooledBytes content) {
        if (!_entries.TryGetValue(name, out int index)) {
            content = default;
            return false;
        }
        content = Read(_records[index]);
        return true;
    }

    /// <summary>The name of the member at a central-directory position.</summary>
    internal string GetName(int index) => DecodeName(_records[index].NameBytes);

    /// <summary>
    /// Reads a member's compressed bytes as stored, for copying it to another archive without inflating it. The local
    /// header is checked against the central directory; the content is not inflated or checked against its CRC-32.
    /// </summary>
    /// <returns><see langword="false"/> when no member has this name, or it is encrypted or uses another method.</returns>
    internal bool TryReadStored(string name, out OfficeZipStoredMember member) {
        member = default;
        if (!_entries.TryGetValue(name, out int index)) return false;
        OfficeArchiveSafety.ZipCentralDirectoryRecord record = _records[index];
        if (!OfficeArchiveSafety.HasSupportedZipFlags(record.Flags, record.CompressionMethod) || record.CompressedSize > int.MaxValue - 64) return false;
        long payload = OfficeArchiveSafety.ValidateEntryPayload(_source, record);
        int count = (int)record.CompressedSize;
        OfficePooledBytes bytes = OfficePooledBytes.Rent(count);
        try {
            if (_buffer != null) Buffer.BlockCopy(_buffer, checked(_bufferOrigin + (int)payload), bytes.Array, 0, count);
            else {
                _source.Position = payload;
                ReadExactly(_source, bytes.Array, count);
            }
        } catch {
            bytes.Dispose();
            throw;
        }
        member = new OfficeZipStoredMember(bytes, record.CompressionMethod, record.Flags, record.DosTime, record.DosDate, record.Crc32, record.UncompressedSize);
        return true;
    }

    private OfficePooledBytes Read(OfficeArchiveSafety.ZipCentralDirectoryRecord record) {
        if (!OfficeArchiveSafety.HasSupportedZipFlags(record.Flags, record.CompressionMethod)) {
            throw new InvalidDataException("A package member is encrypted or uses an unsupported compression method.");
        }
        if (record.UncompressedSize > _maximumEntryLength || record.UncompressedSize > int.MaxValue - 64) {
            throw new InvalidDataException("A package member is larger than the reader accepts.");
        }
        long payload = OfficeArchiveSafety.ValidateEntryPayload(_source, record);
        int length = (int)record.UncompressedSize;
        OfficePooledBytes content = OfficePooledBytes.Rent(length);
        try {
            if (record.CompressionMethod == 0) {
                CopyStored(payload, content.Array, length);
            } else {
                Inflate(payload, record.CompressedSize, content.Array, length);
            }
            if (OfficePngCrc32.Compute(content.Array, 0, length) != record.Crc32) {
                throw new InvalidDataException("A package member does not match its ZIP checksum.");
            }
            return content;
        } catch {
            content.Dispose();
            throw;
        }
    }

    private void CopyStored(long payload, byte[] target, int length) {
        if (_buffer != null) {
            Buffer.BlockCopy(_buffer, checked(_bufferOrigin + (int)payload), target, 0, length);
            return;
        }
        _source.Position = payload;
        ReadExactly(_source, target, length);
    }

    private void Inflate(long payload, long compressedLength, byte[] target, int length) {
        if (compressedLength > int.MaxValue) throw new InvalidDataException("A package member is larger than the reader accepts.");
        int count = (int)compressedLength;
#if NET8_0_OR_GREATER
        // One managed pass decodes the member and proves its Deflate framing ends exactly at the compressed length,
        // which DeflateStream does not report: it stops at the final block and also stops quietly at the end of
        // truncated input.
        if (_buffer != null) {
            if (!OfficeDeflateStreamValidator.TryInflateExact(_buffer, checked(_bufferOrigin + (int)payload), count, target, length, null)) {
                throw new InvalidDataException("A package member has invalid or incomplete Deflate framing.");
            }
            return;
        }
        byte[] compressed = System.Buffers.ArrayPool<byte>.Shared.Rent(Math.Max(1, count));
        try {
            _source.Position = payload;
            ReadExactly(_source, compressed, count);
            if (!OfficeDeflateStreamValidator.TryInflateExact(compressed, 0, count, target, length, null)) {
                throw new InvalidDataException("A package member has invalid or incomplete Deflate framing.");
            }
        } finally {
            System.Buffers.ArrayPool<byte>.Shared.Return(compressed);
        }
#else
        Stream compressed = _buffer != null
            ? new MemoryStream(_buffer, checked(_bufferOrigin + (int)payload), count, writable: false)
            : new SeekableReadWindowStream(_source, payload, compressedLength);
        bool exact = _buffer != null
            ? OfficeDeflateStreamValidator.TryValidateExact(_buffer, checked(_bufferOrigin + (int)payload), count, length)
            : OfficeDeflateStreamValidator.TryValidateExact(compressed, length);
        if (!exact) {
            compressed.Dispose();
            throw new InvalidDataException("A package member has invalid or incomplete Deflate framing.");
        }
        using var inflater = new DeflateStream(compressed, CompressionMode.Decompress, leaveOpen: false);
        ReadExactly(inflater, target, length);
        // A member that inflates past its declared length is corrupt or crafted; reading one more byte proves the end.
        var probe = new byte[1];
        if (inflater.Read(probe, 0, 1) != 0) throw new InvalidDataException("A package member inflates past its declared length.");
#endif
    }

    private static void ReadExactly(Stream stream, byte[] target, int length) {
        int offset = 0;
        while (offset < length) {
            int read = stream.Read(target, offset, length - offset);
            if (read <= 0) throw new InvalidDataException("A package member ends before its declared length.");
            offset += read;
        }
    }

    private static string DecodeName(byte[] bytes) {
        for (int index = 0; index < bytes.Length; index++) {
            if (bytes[index] >= 0x80) {
                // OPC writers percent-encode non-ASCII part names; names that still carry raw bytes are read as UTF-8
                // whether or not the language-encoding flag is set.
                return Encoding.UTF8.GetString(bytes);
            }
        }
#if NET8_0_OR_GREATER
        return Encoding.ASCII.GetString(bytes);
#else
        return Encoding.ASCII.GetString(bytes, 0, bytes.Length);
#endif
    }

    public void Dispose() {
        if (!_leaveOpen) _source.Dispose();
    }
}

/// <summary>A member's compressed bytes and the ZIP fields needed to write it to another archive unchanged.</summary>
internal readonly struct OfficeZipStoredMember : IDisposable {
    internal OfficeZipStoredMember(OfficePooledBytes compressed, ushort method, ushort flags, ushort dosTime, ushort dosDate, uint crc, long size) {
        Compressed = compressed;
        Method = method;
        Flags = flags;
        DosTime = dosTime;
        DosDate = dosDate;
        Crc = crc;
        Size = size;
    }

    internal OfficePooledBytes Compressed { get; }
    internal ushort Method { get; }
    internal ushort Flags { get; }
    internal ushort DosTime { get; }
    internal ushort DosDate { get; }
    internal uint Crc { get; }
    internal long Size { get; }

    public void Dispose() => Compressed.Dispose();
}

/// <summary>A byte buffer rented for one package member; <see cref="Length"/> bytes are valid.</summary>
internal readonly struct OfficePooledBytes : IDisposable {
    private OfficePooledBytes(byte[] array, int length) {
        Array = array;
        Length = length;
    }

    internal byte[] Array { get; }
    internal int Length { get; }

    internal static OfficePooledBytes Rent(int length) {
#if NET8_0_OR_GREATER
        return new OfficePooledBytes(length == 0 ? System.Array.Empty<byte>() : System.Buffers.ArrayPool<byte>.Shared.Rent(length), length);
#else
        return new OfficePooledBytes(length == 0 ? System.Array.Empty<byte>() : new byte[length], length);
#endif
    }

    public void Dispose() {
#if NET8_0_OR_GREATER
        if (Array != null && Array.Length > 0) System.Buffers.ArrayPool<byte>.Shared.Return(Array);
#endif
    }
}

/// <summary>Compares strings ignoring the case of ASCII letters only, as OPC part names compare.</summary>
internal sealed class OfficeAsciiIgnoreCaseComparer : IEqualityComparer<string> {
    internal static readonly OfficeAsciiIgnoreCaseComparer Instance = new();

    public bool Equals(string? x, string? y) {
        if (ReferenceEquals(x, y)) return true;
        if (x == null || y == null || x.Length != y.Length) return false;
        for (int index = 0; index < x.Length; index++) {
            if (Fold(x[index]) != Fold(y[index])) return false;
        }
        return true;
    }

    public int GetHashCode(string value) {
        unchecked {
            int hash = (int)2166136261;
            foreach (char character in value) hash = (hash ^ Fold(character)) * 16777619;
            return hash;
        }
    }

    private static char Fold(char value) => value >= 'A' && value <= 'Z' ? (char)(value + 32) : value;
}
