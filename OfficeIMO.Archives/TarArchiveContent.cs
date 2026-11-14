using System.Globalization;
using System.Text;

namespace OfficeIMO.Archives;

/// <summary>The kind of a TAR member.</summary>
public enum ArchiveEntryKind {
    /// <summary>A regular file with content.</summary>
    File,
    /// <summary>A directory record.</summary>
    Directory,
    /// <summary>A symbolic link. Its target is metadata only; extraction never creates links.</summary>
    SymbolicLink,
    /// <summary>A hard link to an earlier member. Extraction never creates links.</summary>
    HardLink,
    /// <summary>A character device, block device or FIFO record. Extraction never creates these.</summary>
    Special
}

/// <summary>One TAR member's metadata, as authored. Listing never writes files, so names are not path-validated here.</summary>
public sealed class TarArchiveEntry {
    internal TarArchiveEntry(int ordinal, string name, ArchiveEntryKind kind, long length, string? linkTarget,
        DateTimeOffset? lastWriteTime, int mode, long dataOffset) {
        Ordinal = ordinal; Name = name; Kind = kind; Length = length; LinkTarget = linkTarget;
        LastWriteTime = lastWriteTime; Mode = mode; DataOffset = dataOffset;
    }
    /// <summary>Zero-based position among the archive's members, excluding extended headers.</summary>
    public int Ordinal { get; }
    /// <summary>The effective member name after pax, GNU long-name and ustar prefix resolution, without a trailing '/'.</summary>
    public string Name { get; }
    /// <summary>The member kind.</summary>
    public ArchiveEntryKind Kind { get; }
    /// <summary>Content bytes for files; zero for other kinds.</summary>
    public long Length { get; }
    /// <summary>The authored link target for symbolic and hard links; null otherwise. It is never resolved.</summary>
    public string? LinkTarget { get; }
    /// <summary>The member modification time, or null when it falls outside the representable range.</summary>
    public DateTimeOffset? LastWriteTime { get; }
    /// <summary>The Unix permission bits (for example 0644), without the file type.</summary>
    public int Mode { get; }
    internal long DataOffset { get; }
}

/// <summary>Parses POSIX ustar, pax (x/g), GNU (L/K, base-256 numbers) and v7 headers from a seekable TAR stream.</summary>
internal static class TarArchiveContent {
    private const int Block = 512;
    private static readonly UTF8Encoding StrictUtf8 = new(false, true);

    internal static bool IsHeader(byte[] block, int length) {
        if (length < Block) return false;
        if (!TryParseOctal(block, 148, 8, out long stored)) return false;
        if (!ChecksumMatches(block, stored)) return false;
        bool ustar = block[257] == 'u' && block[258] == 's' && block[259] == 't' && block[260] == 'a' && block[261] == 'r';
        return ustar || block[0] != 0;
    }

    internal static bool IsZeroBlock(byte[] block, int offset = 0) {
        for (int index = 0; index < Block; index++) if (block[offset + index] != 0) return false;
        return true;
    }

    /// <summary>Reads every member header, validating checksums, numeric fields, extended headers and that each member's
    /// content lies inside the input. Content is not read.</summary>
    internal static List<TarArchiveEntry> ReadDirectory(Stream source, ArchiveLimits limits, CancellationToken token) {
        var entries = new List<TarArchiveEntry>();
        var global = new Dictionary<string, string>(StringComparer.Ordinal);
        Dictionary<string, string>? local = null;
        string? longName = null, longLink = null;
        long metadata = 0;
        long length = source.Length;
        long position = 0;
        var block = new byte[Block];
        source.Position = 0;
        while (true) {
            token.ThrowIfCancellationRequested();
            bool pending = local != null || longName != null || longLink != null;
            if (position == length) {
                if (pending) throw new InvalidDataException("The TAR archive ends after an extended header.");
                break;
            }
            if (length - position < Block) throw new InvalidDataException("The TAR archive ends inside a header block.");
            ReadExactly(source, position, block);
            if (IsZeroBlock(block)) {
                if (pending) throw new InvalidDataException("An extended TAR header is followed by the end-of-archive marker.");
                // One zero block followed by EOF or a second zero block ends the archive; anything else is corrupt.
                if (position + Block == length) break;
                if (length - position < 2 * Block) throw new InvalidDataException("The TAR end-of-archive marker is truncated.");
                var next = new byte[Block];
                ReadExactly(source, position + Block, next);
                if (!IsZeroBlock(next)) throw new InvalidDataException("A TAR header block is empty.");
                break;
            }
            if (!TryParseOctal(block, 148, 8, out long checksum) || !ChecksumMatches(block, checksum))
                throw new InvalidDataException("A TAR header checksum does not match at offset " + position.ToString(CultureInfo.InvariantCulture) + ".");
            char type = (char)block[156];
            bool posix = block[257] == 'u' && block[258] == 's' && block[259] == 't' && block[260] == 'a' && block[261] == 'r' && block[262] == 0;
            long size = ParseNumber(block, 124, 12, "size");
            if (size < 0) throw new InvalidDataException("A TAR member declares a negative size.");
            long dataOffset = position + Block;
            if (size > long.MaxValue - Block) throw new InvalidDataException("A TAR member size is too large.");
            switch (type) {
                case 'x':
                case 'g': {
                    var records = ReadExtension(source, dataOffset, size, length, limits, ref metadata, token);
                    ParsePax(records, type == 'g' ? global : local ??= new Dictionary<string, string>(StringComparer.Ordinal));
                    position = Advance(dataOffset, size, length);
                    continue;
                }
                case 'L':
                case 'K': {
                    byte[] value = ReadExtension(source, dataOffset, size, length, limits, ref metadata, token);
                    int end = Array.IndexOf(value, (byte)0); if (end < 0) end = value.Length;
                    string text = Decode(value, 0, end, "GNU long " + (type == 'L' ? "name" : "link"));
                    if (type == 'L') longName = text; else longLink = text;
                    position = Advance(dataOffset, size, length);
                    continue;
                }
                case 'V':
                    position = Advance(dataOffset, size, length);
                    continue;
                case 'S': case 'M': case 'N':
                    throw new NotSupportedException("GNU sparse, multi-volume and old long-name TAR members are not supported.");
            }
            Dictionary<string, string>? effective = local == null && global.Count == 0 ? null : Merge(global, local);
            if (effective != null && effective.TryGetValue("size", out string? paxSize)) size = ParseDecimal(paxSize, "size");
            if (effective != null && effective.Keys.Any(key => key.StartsWith("GNU.sparse.", StringComparison.Ordinal)))
                throw new NotSupportedException("Sparse TAR members are not supported.");
            string name = effective != null && effective.TryGetValue("path", out string? paxPath) ? paxPath
                : longName ?? HeaderName(block, posix);
            string? link = effective != null && effective.TryGetValue("linkpath", out string? paxLink) ? paxLink
                : longLink ?? Field(block, 157, 100, "link name");
            ArchiveEntryKind kind;
            switch (type) {
                case '0': case '\0': case '7':
                    kind = name.EndsWith("/", StringComparison.Ordinal) && size == 0 ? ArchiveEntryKind.Directory : ArchiveEntryKind.File; break;
                case '5': case 'D': kind = ArchiveEntryKind.Directory; break;
                case '1': kind = ArchiveEntryKind.HardLink; break;
                case '2': kind = ArchiveEntryKind.SymbolicLink; break;
                case '3': case '4': case '6': kind = ArchiveEntryKind.Special; break;
                default: throw new NotSupportedException("Unsupported TAR member type '" + type + "'.");
            }
            if (dataOffset > length || size > length - dataOffset) throw new InvalidDataException("A TAR member is truncated: " + name);
            DateTimeOffset? time = effective != null && effective.TryGetValue("mtime", out string? paxTime) ? ParseTime(paxTime)
                : ToTime(ParseNumber(block, 136, 12, "mtime"));
            int mode = (int)(ParseNumber(block, 100, 8, "mode") & 0xfff);
            if (entries.Count >= limits.Entries) throw ArchiveLimits.Exceeded(nameof(ArchiveExtractionOptions.MaxEntryCount), limits.Entries);
            limits.Operation.AdmitEntries(1);
            if (kind == ArchiveEntryKind.File) limits.CheckEntry(size);
            entries.Add(new TarArchiveEntry(entries.Count, name.TrimEnd('/'), kind, kind == ArchiveEntryKind.File ? size : 0,
                kind == ArchiveEntryKind.HardLink || kind == ArchiveEntryKind.SymbolicLink ? link : null, time, mode, dataOffset));
            local = null; longName = null; longLink = null;
            position = Advance(dataOffset, size, length);
        }
        return entries;
    }

    /// <summary>Removes leading "./" segments so archives made with "tar -C dir ." map to relative paths. Returns null for the root.</summary>
    internal static string? RelativeName(string name) {
        string value = name;
        while (value.StartsWith("./", StringComparison.Ordinal)) value = value.Substring(2).TrimStart('/');
        return value.Length == 0 || value == "." ? null : value;
    }

    private static long Advance(long dataOffset, long size, long length) {
        long padded = (size + Block - 1) / Block * Block;
        // A producer may omit the last member's padding; the content itself was already bounds-checked.
        return padded > length - dataOffset ? length : dataOffset + padded;
    }

    private static Dictionary<string, string> Merge(Dictionary<string, string> global, Dictionary<string, string>? local) {
        var merged = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var pair in global) if (pair.Value.Length != 0) merged[pair.Key] = pair.Value;
        if (local != null) foreach (var pair in local) {
            if (pair.Value.Length == 0) merged.Remove(pair.Key); else merged[pair.Key] = pair.Value;
        }
        return merged;
    }

    private static byte[] ReadExtension(Stream source, long offset, long size, long length, ArchiveLimits limits, ref long metadata,
        CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (offset > length || size > length - offset) throw new InvalidDataException("A TAR extended header is truncated.");
        ArchiveLimits.Add(ref metadata, size, Math.Min(limits.Metadata, int.MaxValue), nameof(ArchiveExtractionOptions.MaxMetadataBytes));
        var value = new byte[size];
        ReadExactly(source, offset, value);
        return value;
    }

    private static void ParsePax(byte[] data, Dictionary<string, string> target) {
        int offset = 0;
        var records = new Dictionary<string, string>(StringComparer.Ordinal);
        while (offset < data.Length) {
            if (data[offset] == 0) {
                for (int index = offset; index < data.Length; index++)
                    if (data[index] != 0) throw new InvalidDataException("A pax header has data after its records.");
                break;
            }
            int space = Array.IndexOf(data, (byte)' ', offset);
            if (space <= offset || space - offset > 19) throw new InvalidDataException("A pax record length is malformed.");
            long recordLength = 0;
            for (int index = offset; index < space; index++) {
                if (data[index] < '0' || data[index] > '9') throw new InvalidDataException("A pax record length is malformed.");
                recordLength = recordLength * 10 + (data[index] - '0');
            }
            if (recordLength < space - offset + 3 || recordLength > data.Length - offset || data[offset + recordLength - 1] != '\n')
                throw new InvalidDataException("A pax record is malformed.");
            int end = offset + (int)recordLength - 1;
            int equals = Array.IndexOf(data, (byte)'=', space + 1, end - space - 1);
            if (equals <= space + 1) throw new InvalidDataException("A pax record has no key.");
            string key = Decode(data, space + 1, equals - space - 1, "pax key");
            string value = Decode(data, equals + 1, end - equals - 1, "pax value");
            records[key] = value;
            offset = end + 1;
        }
        if (records.TryGetValue("hdrcharset", out string? charset) && charset != "ISO-IR 10646 2000 UTF-8")
            throw new NotSupportedException("pax headers with a non-UTF-8 character set are not supported.");
        foreach (var pair in records) target[pair.Key] = pair.Value;
    }

    private static string HeaderName(byte[] block, bool posix) {
        string name = Field(block, 0, 100, "name");
        if (!posix || block[345] == 0) return name;
        return Field(block, 345, 155, "prefix") + "/" + name;
    }

    private static string Field(byte[] block, int offset, int length, string field) {
        int end = offset;
        while (end < offset + length && block[end] != 0) end++;
        return Decode(block, offset, end - offset, field);
    }

    private static string Decode(byte[] bytes, int offset, int count, string field) {
        try { return StrictUtf8.GetString(bytes, offset, count); }
        catch (DecoderFallbackException error) { throw new InvalidDataException("The TAR " + field + " is not valid UTF-8.", error); }
    }

    private static bool ChecksumMatches(byte[] block, long stored) {
        long unsignedSum = 0, signedSum = 0;
        for (int index = 0; index < Block; index++) {
            byte value = index >= 148 && index < 156 ? (byte)' ' : block[index];
            unsignedSum += value; signedSum += (sbyte)value;
        }
        return stored == unsignedSum || stored == signedSum;
    }

    internal static long ParseNumber(byte[] block, int offset, int length, string field) {
        if ((block[offset] & 0x80) != 0) {
            // GNU base-256: big-endian two's complement with the marker bit in the first byte.
            bool negative = (block[offset] & 0x40) != 0;
            long value = negative ? -1 : 0;
            for (int index = 0; index < length; index++) {
                byte part = index == 0 ? (byte)(negative ? block[offset] | 0x80 : block[offset] & 0x7f) : block[offset + index];
                if (index < length - 8 && part != (negative ? 0xff : 0))
                    throw new InvalidDataException("The TAR " + field + " field overflows.");
                value = value << 8 | part;
            }
            if (!negative && value < 0) throw new InvalidDataException("The TAR " + field + " field overflows.");
            return value;
        }
        if (!TryParseOctal(block, offset, length, out long result)) throw new InvalidDataException("The TAR " + field + " field is not octal.");
        return result;
    }

    private static bool TryParseOctal(byte[] block, int offset, int length, out long value) {
        value = 0;
        int index = offset, end = offset + length;
        while (index < end && (block[index] == ' ' || block[index] == 0)) index++;
        for (; index < end && block[index] >= '0' && block[index] <= '7'; index++) {
            if (value > long.MaxValue >> 3) return false;
            value = value << 3 | (long)(block[index] - '0');
        }
        for (; index < end; index++) if (block[index] != ' ' && block[index] != 0) return false;
        return true;
    }

    private static long ParseDecimal(string text, string field) {
        if (text.Length == 0 || text.Length > 19 || text.Any(character => character < '0' || character > '9')
            || !long.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out long value))
            throw new InvalidDataException("The pax " + field + " value is not a nonnegative decimal number.");
        return value;
    }

    private static DateTimeOffset? ParseTime(string text) {
        int dot = text.IndexOf('.');
        string whole = dot < 0 ? text : text.Substring(0, dot);
        bool negative = whole.StartsWith("-", StringComparison.Ordinal);
        string digits = negative ? whole.Substring(1) : whole;
        string fraction = dot < 0 ? string.Empty : text.Substring(dot + 1);
        if (digits.Length == 0 || digits.Length > 18 || digits.Any(character => character < '0' || character > '9')
            || fraction.Any(character => character < '0' || character > '9'))
            throw new InvalidDataException("The pax mtime value is malformed.");
        long seconds = long.Parse(digits, NumberStyles.None, CultureInfo.InvariantCulture);
        DateTimeOffset? time = ToTime(negative ? -seconds : seconds);
        if (time == null || fraction.Length == 0) return time;
        long ticks = long.Parse((fraction + "0000000").Substring(0, 7), NumberStyles.None, CultureInfo.InvariantCulture);
        try { return negative ? time.Value.AddTicks(-ticks) : time.Value.AddTicks(ticks); }
        catch (ArgumentOutOfRangeException) { return null; }
    }

    private static DateTimeOffset? ToTime(long seconds) =>
        seconds < -62135596800L || seconds > 253402300799L ? null : DateTimeOffset.FromUnixTimeSeconds(seconds);

    internal static void ReadExactly(Stream source, long offset, byte[] buffer) {
        source.Position = offset;
        int total = 0, read;
        while (total < buffer.Length && (read = source.Read(buffer, total, buffer.Length - total)) > 0) total += read;
        if (total != buffer.Length) throw new InvalidDataException("The TAR archive is truncated.");
    }
}
