using System.IO.Compression;
using System.IO.Packaging;
using System.Xml;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Word;

/// <summary>
/// An editable OPC package over a ZIP archive held in a memory stream the document owns.
/// </summary>
/// <remarks>
/// <para>System.IO.Packaging rewrites the whole archive on every flush: on .NET it closes the ZIP archive, which
/// recompresses every entry it ever opened, and the Open XML SDK then reopens the package; on .NET Framework the
/// document saved a fresh copy that recompressed every part. This package keeps the archive's compressed entries and
/// writes a part again only after something wrote to it. A flush copies every other entry's compressed bytes
/// unchanged, so untouched parts stay byte-identical.</para>
/// <para>Relationships, core properties and part bookkeeping come from <see cref="Package"/>, as for ZipPackage.
/// Content types follow ZipPackage's rules for new and removed parts. <see cref="TryOpen"/> accepts only archives whose
/// every entry maps to one part with a content type, which covers documents written by Word and the Open XML SDK; any
/// other layout returns <see langword="null"/> so the caller keeps the System.IO.Packaging route.</para>
/// </remarks>
internal sealed class WordZipPackage : Package {
    private const string ContentTypesEntryName = "[Content_Types].xml";
    private const string ContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";
    private const int MaximumEntries = ushort.MaxValue - 1;
    private const ushort Utf8NameFlag = 0x0800;
    private const ushort DataDescriptorFlag = 0x0008;
    private static readonly byte[] Utf8Preamble = { 0xEF, 0xBB, 0xBF };
    private static readonly char[] AttributeSpecialCharacters = { '&', '<', '>', '"' };

    private readonly MemoryStream _container;
    // Entries in archive order. New parts are appended, as ZipArchive appends new entries.
    private readonly List<Entry> _entries;
    private readonly Dictionary<string, Part> _parts = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _defaults;
    private readonly Dictionary<string, KeyValuePair<string, string>> _overrides;
    private readonly Entry _contentTypes;
    private bool _structureChanged;
    private bool _contentTypesChanged;
    private bool _abandoned;

    private WordZipPackage(MemoryStream container, List<Entry> entries, Entry contentTypes,
        Dictionary<string, string> defaults, Dictionary<string, KeyValuePair<string, string>> overrides)
        : base(FileAccess.ReadWrite) {
        _container = container;
        _entries = entries;
        _contentTypes = contentTypes;
        _defaults = defaults;
        _overrides = overrides;
        foreach (Entry entry in entries) {
            if (entry.PartUri == null) continue;
            var part = new Part(this, entry.PartUri, entry.ContentType!, entry);
            entry.Part = part;
            _parts.Add(Key(entry.PartUri), part);
        }
    }

    /// <summary>
    /// Opens the archive in <paramref name="container"/> for editing, or returns <see langword="null"/> when its layout
    /// needs System.IO.Packaging (ZIP64, encryption, unmapped or interleaved entries, unusual names or content types).
    /// </summary>
    /// <remarks>The stream is read but not written until the package is flushed.</remarks>
    internal static WordZipPackage? TryOpen(MemoryStream container) {
        if (!container.CanRead || !container.CanSeek || !container.CanWrite || !container.TryGetBuffer(out _)) return null;
        try {
            // Flushing rewrites the archive in place, which a fixed-size stream cannot hold.
            container.Capacity = container.Capacity;
        } catch (NotSupportedException) {
            return null;
        }
        long originalPosition = container.Position;
        try {
            container.Position = 0;
            var records = new List<OfficeArchiveSafety.ZipCentralDirectoryRecord>();
            OfficeArchiveSafety.ZipCentralDirectoryScanResult scan =
                OfficeArchiveSafety.ScanZipCentralDirectoryForExtraction(container, container.Length, MaximumEntries, records);
            if (!scan.IsValid || scan.LimitExceeded) return null;

            var entries = new List<Entry>(records.Count);
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Entry? contentTypes = null;
            foreach (OfficeArchiveSafety.ZipCentralDirectoryRecord record in records) {
                if (record.UsesZip64 || record.UncompressedSize > int.MaxValue || record.CompressedSize > int.MaxValue) return null;
                if (!OfficeArchiveSafety.HasSupportedZipFlags(record.Flags, record.CompressionMethod)) return null;
                string? name = DecodeName(record.NameBytes, record.Flags);
                if (name == null || !names.Add(name)) return null;
                long payload = OfficeArchiveSafety.ValidateEntryPayload(container, record);
                var entry = new Entry(name, record.NameBytes, record) { PayloadOffset = payload };
                if (string.Equals(name, ContentTypesEntryName, StringComparison.OrdinalIgnoreCase)) {
                    if (!string.Equals(name, ContentTypesEntryName, StringComparison.Ordinal)) return null;
                    contentTypes = entry;
                } else {
                    entry.PartUri = TryCreatePartUri(name);
                    if (entry.PartUri == null) return null;
                }
                entries.Add(entry);
            }
            if (contentTypes == null) return null;

            byte[] contentTypesXml = Inflate(container, contentTypes, out int contentTypesLength);
            if (!TryReadContentTypes(contentTypesXml, contentTypesLength,
                    out Dictionary<string, string> defaults, out Dictionary<string, KeyValuePair<string, string>> overrides)) {
                return null;
            }
            foreach (Entry entry in entries) {
                if (entry.PartUri == null) continue;
                entry.ContentType = ResolveContentType(entry.PartUri, defaults, overrides);
                if (entry.ContentType == null) return null;
            }
            var package = new WordZipPackage(container, entries, contentTypes, defaults, overrides);
            // Package.Open loads the part list up front, and PartExists answers only from that list.
            package.GetParts();
            return package;
        } catch (InvalidOperationException) {
            return null;
        } catch (InvalidDataException) {
            return null;
        } catch (XmlException) {
            return null;
        } catch (FormatException) {
            return null;
        } catch (ArgumentException) {
            return null;
        } finally {
            container.Position = originalPosition;
        }
    }

    /// <summary>Stops this package from writing to its container, so the caller can open the same bytes another way.</summary>
    internal void Abandon() => _abandoned = true;

    protected override PackagePart CreatePartCore(Uri partUri, string contentType, CompressionOption compressionOption) {
        string name = partUri.OriginalString.Substring(1);
        byte[] nameBytes = System.Text.Encoding.UTF8.GetBytes(name);
        var entry = new Entry(name, nameBytes, flags: IsAscii(name) ? (ushort)0 : Utf8NameFlag) {
            PartUri = partUri,
            ContentType = contentType
        };
        var part = new Part(this, partUri, contentType, entry);
        entry.Part = part;
        _entries.Add(entry);
        _parts[Key(partUri)] = part;
        AddContentType(partUri, contentType);
        _structureChanged = true;
        return part;
    }

    protected override PackagePart? GetPartCore(Uri partUri) =>
        _parts.TryGetValue(Key(partUri), out Part? part) ? part : null;

    protected override void DeletePartCore(Uri partUri) {
        string key = Key(partUri);
        if (!_parts.TryGetValue(key, out Part? part)) return;
        _parts.Remove(key);
        _entries.Remove(part.Entry);
        if (_overrides.Remove(key)) _contentTypesChanged = true;
        _structureChanged = true;
    }

    protected override PackagePart[] GetPartsCore() {
        var parts = new List<PackagePart>(_parts.Count);
        foreach (Entry entry in _entries) {
            if (entry.Part != null) parts.Add(entry.Part);
        }
        return parts.ToArray();
    }

    /// <summary>Writes the archive into the container when a part, the part set or the content types changed.</summary>
    protected override void FlushCore() {
        if (_abandoned || !HasChanges()) return;

        byte[]? contentTypesXml = _contentTypesChanged ? SerializeContentTypes() : null;
        // Copied entries keep their compressed size; rewritten XML usually deflates to well under a quarter, and the
        // stream grows if a part compresses worse. Reserving the uncompressed size would over-allocate large documents.
        long capacity = 22;
        foreach (Entry entry in _entries) {
            capacity += 76 + 2L * entry.NameBytes.Length;
            if (ReferenceEquals(entry, _contentTypes)) {
                capacity += contentTypesXml?.Length ?? entry.CompressedSize;
            } else {
                capacity += entry.Part!.IsUnchanged ? entry.CompressedSize : entry.Part.Length / 4 + 64;
            }
        }
        capacity = Math.Min(capacity, int.MaxValue);

        // The new archive is written over the old one, so the bytes it copies from are set aside first.
        int previousLength = checked((int)_container.Length);
        ArraySegment<byte> current = GetContainerBuffer();
        byte[] previous = RentBuffer(previousLength);
        Buffer.BlockCopy(current.Array!, current.Offset, previous, 0, previousLength);
        var source = new ArraySegment<byte>(previous, 0, previousLength);
        var written = new List<WrittenEntry>(_entries.Count);
        try {
            _container.SetLength(0);
            if (_container.Capacity < capacity) _container.Capacity = (int)capacity;
            WriteArchive(_container, source, contentTypesXml, written);
        } catch {
            _container.SetLength(0);
            _container.Write(previous, 0, previousLength);
            _container.Position = 0;
            throw;
        } finally {
            ReturnBuffer(previous);
        }
        _container.Position = 0;
        foreach (WrittenEntry entry in written) entry.Entry.Rebase(entry);
        _structureChanged = false;
        _contentTypesChanged = false;
    }

    private void WriteArchive(MemoryStream output, ArraySegment<byte> source, byte[]? contentTypesXml, List<WrittenEntry> written) {
        DateTime now = DateTime.Now;
        foreach (Entry entry in _entries) {
            long headerOffset = output.Position;
            if (ReferenceEquals(entry, _contentTypes) ? contentTypesXml == null : entry.Part!.IsUnchanged) {
                WriteCopiedEntry(output, entry, source);
                written.Add(new WrittenEntry(entry, headerOffset, entry.Crc32, entry.CompressedSize, entry.UncompressedSize,
                    entry.Method, entry.Flags, entry.DosTime, entry.DosDate));
                continue;
            }

            byte[] data;
            int length;
            if (ReferenceEquals(entry, _contentTypes)) {
                data = contentTypesXml!;
                length = data.Length;
            } else {
                data = entry.Part!.Data;
                length = entry.Part.Length;
            }
            if (!entry.HasArchiveData) {
                ToDosDateTime(now, out ushort time, out ushort date);
                entry.DosTime = time;
                entry.DosDate = date;
            }
            written.Add(WriteNewEntry(output, entry, data, length, headerOffset));
        }

        long centralDirectoryOffset = output.Position;
        foreach (WrittenEntry entry in written) WriteCentralDirectoryHeader(output, entry);
        long centralDirectorySize = output.Position - centralDirectoryOffset;
        if (written.Count > MaximumEntries || centralDirectoryOffset > uint.MaxValue || centralDirectorySize > uint.MaxValue) {
            throw new IOException("The Word package needs ZIP64, which this package writer does not produce.");
        }
        WriteEndOfCentralDirectory(output, written.Count, centralDirectorySize, centralDirectoryOffset);
    }

    private static byte[] RentBuffer(int length) =>
#if NET8_0_OR_GREATER
        System.Buffers.ArrayPool<byte>.Shared.Rent(length);
#else
        new byte[length];
#endif

    private static void ReturnBuffer(byte[] buffer) {
#if NET8_0_OR_GREATER
        System.Buffers.ArrayPool<byte>.Shared.Return(buffer);
#endif
    }

    private bool HasChanges() {
        if (_structureChanged || _contentTypesChanged) return true;
        foreach (Part part in _parts.Values) {
            if (!part.IsUnchanged) return true;
        }
        return false;
    }

    private ArraySegment<byte> GetContainerBuffer() {
        if (_container.TryGetBuffer(out ArraySegment<byte> buffer)) return buffer;
        throw new InvalidOperationException("The Word package stream no longer exposes its buffer.");
    }

    private byte[] ReadEntry(Entry entry, out int length) => Inflate(_container, entry, out length);

    private static byte[] Inflate(MemoryStream container, Entry entry, out int length) {
        if (!container.TryGetBuffer(out ArraySegment<byte> buffer)) {
            throw new InvalidOperationException("The Word package stream no longer exposes its buffer.");
        }
        int compressedSize = checked((int)entry.CompressedSize);
        int offset = checked(buffer.Offset + (int)entry.PayloadOffset);
        int expected = checked((int)entry.UncompressedSize);
        if (entry.Method == 0) {
            var stored = new byte[compressedSize];
            Buffer.BlockCopy(buffer.Array!, offset, stored, 0, compressedSize);
            length = compressedSize;
            return stored;
        }

        var data = new byte[Math.Max(expected, 16)];
        length = 0;
        using var input = new MemoryStream(buffer.Array!, offset, compressedSize, writable: false);
        using var deflate = new DeflateStream(input, CompressionMode.Decompress);
        while (true) {
            if (length == data.Length) {
                // The declared size is full; grow only if the entry really holds more.
                int next = deflate.ReadByte();
                if (next < 0) break;
                Array.Resize(ref data, checked(data.Length * 2));
                data[length++] = (byte)next;
            }
            int read = deflate.Read(data, length, data.Length - length);
            if (read == 0) break;
            length += read;
        }
        return data;
    }

    private static void WriteCopiedEntry(MemoryStream output, Entry entry, ArraySegment<byte> source) {
        WriteLocalHeader(output, entry.NameBytes, entry.Flags, entry.Method, entry.DosTime, entry.DosDate,
            entry.Crc32, entry.CompressedSize, entry.UncompressedSize);
        output.Write(source.Array!, checked(source.Offset + (int)entry.PayloadOffset), checked((int)entry.CompressedSize));
    }

    private static WrittenEntry WriteNewEntry(MemoryStream output, Entry entry, byte[] data, int length, long headerOffset) {
        uint crc = OfficePngCrc32.Compute(data, 0, length);
        ushort flags = (ushort)(entry.Flags & Utf8NameFlag);
        // Placeholder sizes are patched once the compressed length is known.
        WriteLocalHeader(output, entry.NameBytes, flags, 8, entry.DosTime, entry.DosDate, crc, 0, length);
        long dataOffset = output.Position;
        ushort method = 8;
        if (length > 0) {
            using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true)) {
                deflate.Write(data, 0, length);
            }
        }
        long compressed = output.Position - dataOffset;
        if (length == 0 || compressed >= length) {
            output.SetLength(dataOffset);
            output.Position = dataOffset;
            output.Write(data, 0, length);
            method = 0;
            compressed = length;
        }
        long end = output.Position;
        output.Position = headerOffset + 8;
        WriteUInt16(output, method);
        output.Position = headerOffset + 18;
        WriteUInt32(output, checked((uint)compressed));
        output.Position = end;
        return new WrittenEntry(entry, headerOffset, crc, compressed, length, method, flags, entry.DosTime, entry.DosDate);
    }

    private static void WriteLocalHeader(Stream output, byte[] name, ushort flags, ushort method, ushort time, ushort date,
        uint crc, long compressedSize, long uncompressedSize) {
        WriteUInt32(output, 0x04034b50);
        WriteUInt16(output, 20);
        WriteUInt16(output, (ushort)(flags & ~DataDescriptorFlag));
        WriteUInt16(output, method);
        WriteUInt16(output, time);
        WriteUInt16(output, date);
        WriteUInt32(output, crc);
        WriteUInt32(output, checked((uint)compressedSize));
        WriteUInt32(output, checked((uint)uncompressedSize));
        WriteUInt16(output, checked((ushort)name.Length));
        WriteUInt16(output, 0);
        output.Write(name, 0, name.Length);
    }

    private static void WriteCentralDirectoryHeader(Stream output, WrittenEntry entry) {
        byte[] name = entry.Entry.NameBytes;
        WriteUInt32(output, 0x02014b50);
        WriteUInt16(output, 20);
        WriteUInt16(output, 20);
        WriteUInt16(output, (ushort)(entry.Flags & ~DataDescriptorFlag));
        WriteUInt16(output, entry.Method);
        WriteUInt16(output, entry.DosTime);
        WriteUInt16(output, entry.DosDate);
        WriteUInt32(output, entry.Crc32);
        WriteUInt32(output, checked((uint)entry.CompressedSize));
        WriteUInt32(output, checked((uint)entry.UncompressedSize));
        WriteUInt16(output, checked((ushort)name.Length));
        WriteUInt16(output, 0);
        WriteUInt16(output, 0);
        WriteUInt16(output, 0);
        WriteUInt16(output, 0);
        WriteUInt32(output, 0);
        WriteUInt32(output, checked((uint)entry.HeaderOffset));
        output.Write(name, 0, name.Length);
    }

    private static void WriteEndOfCentralDirectory(Stream output, int count, long size, long offset) {
        WriteUInt32(output, 0x06054b50);
        WriteUInt16(output, 0);
        WriteUInt16(output, 0);
        WriteUInt16(output, (ushort)count);
        WriteUInt16(output, (ushort)count);
        WriteUInt32(output, (uint)size);
        WriteUInt32(output, (uint)offset);
        WriteUInt16(output, 0);
    }

    private static void WriteUInt16(Stream output, ushort value) {
        output.WriteByte((byte)value);
        output.WriteByte((byte)(value >> 8));
    }

    private static void WriteUInt32(Stream output, uint value) {
        output.WriteByte((byte)value);
        output.WriteByte((byte)(value >> 8));
        output.WriteByte((byte)(value >> 16));
        output.WriteByte((byte)(value >> 24));
    }

    private static void ToDosDateTime(DateTime value, out ushort time, out ushort date) {
        if (value.Year < 1980) value = new DateTime(1980, 1, 1);
        time = (ushort)((value.Hour << 11) | (value.Minute << 5) | (value.Second / 2));
        date = (ushort)(((value.Year - 1980) << 9) | (value.Month << 5) | value.Day);
    }

    // ZipPackage's mapping: a new extension becomes a default; an extension whose default differs needs an override.
    private void AddContentType(Uri partUri, string contentType) {
        string extension = GetExtension(partUri.OriginalString);
        if (extension.Length == 0 ||
            (_defaults.TryGetValue(extension, out string? existing) && !AreSameMediaType(existing, contentType))) {
            _overrides[Key(partUri)] = new KeyValuePair<string, string>(partUri.OriginalString, contentType);
            _contentTypesChanged = true;
        } else if (existing == null) {
            _defaults.Add(extension, contentType);
            _contentTypesChanged = true;
        }
    }

    private byte[] SerializeContentTypes() {
        var builder = new System.Text.StringBuilder(256 + 128 * (_defaults.Count + _overrides.Count));
        builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?><Types xmlns=\"").Append(ContentTypesNamespace).Append("\">");
        foreach (KeyValuePair<string, string> item in _defaults) {
            builder.Append("<Default Extension=\"").Append(EscapeAttribute(item.Key))
                .Append("\" ContentType=\"").Append(EscapeAttribute(item.Value)).Append("\" />");
        }
        foreach (KeyValuePair<string, string> item in _overrides.Values) {
            builder.Append("<Override PartName=\"").Append(EscapeAttribute(item.Key))
                .Append("\" ContentType=\"").Append(EscapeAttribute(item.Value)).Append("\" />");
        }
        builder.Append("</Types>");
        string xml = builder.ToString();
        var bytes = new byte[Utf8Preamble.Length + System.Text.Encoding.UTF8.GetByteCount(xml)];
        Buffer.BlockCopy(Utf8Preamble, 0, bytes, 0, Utf8Preamble.Length);
        System.Text.Encoding.UTF8.GetBytes(xml, 0, xml.Length, bytes, Utf8Preamble.Length);
        return bytes;
    }

    private static bool TryReadContentTypes(byte[] xml, int length,
        out Dictionary<string, string> defaults, out Dictionary<string, KeyValuePair<string, string>> overrides) {
        defaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        overrides = new Dictionary<string, KeyValuePair<string, string>>(StringComparer.OrdinalIgnoreCase);
        using var input = new MemoryStream(xml, 0, length, writable: false);
        using XmlReader reader = XmlReader.Create(input, new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = true,
            IgnoreProcessingInstructions = true
        });
        if (reader.MoveToContent() != XmlNodeType.Element || reader.LocalName != "Types" ||
            reader.NamespaceURI != ContentTypesNamespace || reader.HasAttributes && !OnlyNamespaceAttributes(reader)) {
            return false;
        }
        if (reader.IsEmptyElement) return true;
        reader.Read();
        while (!reader.EOF) {
            if (reader.NodeType == XmlNodeType.EndElement) {
                reader.Read();
                return reader.EOF || reader.MoveToContent() == XmlNodeType.None;
            }
            if (reader.NodeType != XmlNodeType.Element || reader.NamespaceURI != ContentTypesNamespace ||
                !reader.IsEmptyElement) {
                return false;
            }
            string? contentType = reader.GetAttribute("ContentType");
            if (string.IsNullOrEmpty(contentType) || reader.AttributeCount != 2) return false;
            if (reader.LocalName == "Default") {
                string? extension = reader.GetAttribute("Extension");
                if (string.IsNullOrEmpty(extension) || defaults.ContainsKey(extension!)) return false;
                defaults.Add(extension!, contentType!);
            } else if (reader.LocalName == "Override") {
                string? partName = reader.GetAttribute("PartName");
                if (string.IsNullOrEmpty(partName)) return false;
                Uri? partUri = TryCreatePartUri(partName!.Substring(partName.StartsWith("/", StringComparison.Ordinal) ? 1 : 0));
                if (partUri == null || !string.Equals(partUri.OriginalString, partName, StringComparison.Ordinal)) return false;
                string key = Key(partUri);
                if (overrides.ContainsKey(key)) return false;
                overrides.Add(key, new KeyValuePair<string, string>(partName, contentType!));
            } else {
                return false;
            }
            reader.Read();
        }
        return false;
    }

    private static bool OnlyNamespaceAttributes(XmlReader reader) {
        bool onlyNamespaces = true;
        while (reader.MoveToNextAttribute()) {
            if (reader.Prefix != "xmlns" && reader.LocalName != "xmlns") onlyNamespaces = false;
        }
        reader.MoveToElement();
        return onlyNamespaces;
    }

    private static string? ResolveContentType(Uri partUri, Dictionary<string, string> defaults,
        Dictionary<string, KeyValuePair<string, string>> overrides) {
        if (overrides.TryGetValue(Key(partUri), out KeyValuePair<string, string> item)) return item.Value;
        string extension = GetExtension(partUri.OriginalString);
        return extension.Length > 0 && defaults.TryGetValue(extension, out string? contentType) ? contentType : null;
    }

    private static string GetExtension(string partName) {
        int slash = partName.LastIndexOf('/');
        int dot = partName.LastIndexOf('.');
        return dot > slash && dot < partName.Length - 1 ? partName.Substring(dot + 1) : string.Empty;
    }

    private static bool AreSameMediaType(string left, string right) =>
        string.Equals(MediaType(left), MediaType(right), StringComparison.OrdinalIgnoreCase);

    private static string MediaType(string contentType) {
        int parameters = contentType.IndexOf(';');
        return (parameters < 0 ? contentType : contentType.Substring(0, parameters)).Trim();
    }

    private static Uri? TryCreatePartUri(string entryName) {
        if (entryName.Length == 0 || entryName.EndsWith("/", StringComparison.Ordinal) ||
            entryName.StartsWith("/", StringComparison.Ordinal) || entryName.IndexOf('\\') >= 0 ||
            entryName.IndexOf('%') >= 0) {
            return null;
        }
        for (int index = 0; index < entryName.Length; index++) {
            char value = entryName[index];
            if (value <= ' ' || value >= 0x7F || value == '[' || value == ']' || value == '#' || value == '?') return null;
        }
        try {
            Uri partUri = PackUriHelper.CreatePartUri(new Uri("/" + entryName, UriKind.Relative));
            return string.Equals(partUri.OriginalString, "/" + entryName, StringComparison.Ordinal) ? partUri : null;
        } catch (ArgumentException) {
            return null;
        } catch (UriFormatException) {
            return null;
        }
    }

    private static string? DecodeName(byte[] nameBytes, ushort flags) {
        bool ascii = true;
        foreach (byte value in nameBytes) {
            if (value >= 0x80) ascii = false;
        }
        if (ascii) return System.Text.Encoding.ASCII.GetString(nameBytes);
        // Non-ASCII names only appear as content types; those parts take the System.IO.Packaging route anyway.
        return (flags & Utf8NameFlag) != 0 ? System.Text.Encoding.UTF8.GetString(nameBytes) : null;
    }

    private static bool IsAscii(string value) {
        foreach (char character in value) {
            if (character >= 0x80) return false;
        }
        return true;
    }

    private static string EscapeAttribute(string value) {
        if (value.IndexOfAny(AttributeSpecialCharacters) < 0) return value;
        return value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
    }

    private static string Key(Uri partUri) => partUri.OriginalString;

    /// <summary>One archive entry: a part, or the content types catalog.</summary>
    private sealed class Entry {
        internal Entry(string name, byte[] nameBytes, OfficeArchiveSafety.ZipCentralDirectoryRecord record) {
            Name = name;
            NameBytes = nameBytes;
            Flags = record.Flags;
            Method = record.CompressionMethod;
            Crc32 = record.Crc32;
            CompressedSize = record.CompressedSize;
            UncompressedSize = record.UncompressedSize;
            DosTime = record.DosTime;
            DosDate = record.DosDate;
            HasArchiveData = true;
        }

        internal Entry(string name, byte[] nameBytes, ushort flags) {
            Name = name;
            NameBytes = nameBytes;
            Flags = flags;
            PayloadOffset = -1;
        }

        internal string Name { get; }
        internal byte[] NameBytes { get; }
        internal Uri? PartUri { get; set; }
        internal string? ContentType { get; set; }
        internal Part? Part { get; set; }
        internal ushort Flags { get; private set; }
        internal ushort Method { get; private set; }
        internal uint Crc32 { get; private set; }
        internal long CompressedSize { get; private set; }
        internal long UncompressedSize { get; private set; }
        internal ushort DosTime { get; set; }
        internal ushort DosDate { get; set; }
        internal long PayloadOffset { get; set; }
        /// <summary>Whether the container holds this entry's compressed bytes.</summary>
        internal bool HasArchiveData { get; private set; }

        internal void Rebase(WrittenEntry written) {
            Flags = written.Flags;
            Method = written.Method;
            Crc32 = written.Crc32;
            CompressedSize = written.CompressedSize;
            UncompressedSize = written.UncompressedSize;
            PayloadOffset = written.HeaderOffset + 30 + NameBytes.Length;
            HasArchiveData = true;
            Part?.MarkWritten();
        }
    }

    private readonly struct WrittenEntry {
        internal WrittenEntry(Entry entry, long headerOffset, uint crc32, long compressedSize, long uncompressedSize,
            ushort method, ushort flags, ushort dosTime, ushort dosDate) {
            Entry = entry;
            HeaderOffset = headerOffset;
            Crc32 = crc32;
            CompressedSize = compressedSize;
            UncompressedSize = uncompressedSize;
            Method = method;
            Flags = flags;
            DosTime = dosTime;
            DosDate = dosDate;
        }

        internal Entry Entry { get; }
        internal long HeaderOffset { get; }
        internal uint Crc32 { get; }
        internal long CompressedSize { get; }
        internal long UncompressedSize { get; }
        internal ushort Method { get; }
        internal ushort Flags { get; }
        internal ushort DosTime { get; }
        internal ushort DosDate { get; }
    }

    /// <summary>A part whose bytes are inflated on first use and written again only after a change.</summary>
    private sealed class Part : PackagePart {
        private readonly WordZipPackage _package;
        private byte[]? _data;
        private int _length;
        private bool _changed;

        internal Part(WordZipPackage package, Uri partUri, string contentType, Entry entry)
            : base(package, partUri, contentType, CompressionOption.Normal) {
            _package = package;
            Entry = entry;
            _changed = !entry.HasArchiveData;
        }

        internal Entry Entry { get; }
        internal bool IsUnchanged => !_changed && Entry.HasArchiveData;
        internal int Length => _length;

        internal byte[] Data {
            get {
                EnsureData();
                return _data!;
            }
        }

        internal void MarkWritten() => _changed = false;

        protected override Stream GetStreamCore(FileMode mode, FileAccess access) {
            if (mode == FileMode.Create || mode == FileMode.Truncate) {
                _data ??= new byte[256];
                _length = 0;
                _changed = true;
            } else {
                EnsureData();
            }
            var stream = new PartStream(this, access != FileAccess.Read);
            if (mode == FileMode.Append) stream.Position = _length;
            return stream;
        }

        private void EnsureData() {
            if (_data != null) return;
            if (Entry.HasArchiveData) {
                _data = _package.ReadEntry(Entry, out _length);
            } else {
                _data = new byte[256];
                _length = 0;
            }
        }

        internal int Read(long position, byte[] buffer, int offset, int count) {
            if (position >= _length) return 0;
            int available = (int)Math.Min(count, _length - position);
            Buffer.BlockCopy(_data!, (int)position, buffer, offset, available);
            return available;
        }

        internal void Write(long position, byte[] buffer, int offset, int count) {
            long end = checked(position + count);
            if (end > int.MaxValue) throw new IOException("A Word package part cannot exceed 2 GB.");
            EnsureCapacity((int)end);
            if (position > _length) Array.Clear(_data!, _length, (int)position - _length);
            Buffer.BlockCopy(buffer, offset, _data!, (int)position, count);
            if (end > _length) _length = (int)end;
            _changed = true;
        }

        internal void SetLength(long value) {
            if (value > int.MaxValue) throw new IOException("A Word package part cannot exceed 2 GB.");
            int length = (int)value;
            EnsureCapacity(length);
            if (length > _length) Array.Clear(_data!, _length, length - _length);
            _length = length;
            _changed = true;
        }

        private void EnsureCapacity(int required) {
            if (required <= _data!.Length) return;
            int capacity = Math.Max(required, (int)Math.Min(int.MaxValue, _data.Length * 2L));
            Array.Resize(ref _data, capacity);
        }
    }

    /// <summary>A read or read-write view of a part's bytes with its own position.</summary>
    private sealed class PartStream : Stream {
        private readonly Part _part;
        private readonly bool _writable;
        private long _position;
        private bool _disposed;

        internal PartStream(Part part, bool writable) {
            _part = part;
            _writable = writable;
        }

        public override bool CanRead => !_disposed;
        public override bool CanSeek => !_disposed;
        public override bool CanWrite => !_disposed && _writable;
        public override long Length {
            get {
                ThrowIfDisposed();
                return _part.Length;
            }
        }

        public override long Position {
            get {
                ThrowIfDisposed();
                return _position;
            }
            set {
                ThrowIfDisposed();
                if (value < 0) throw new ArgumentOutOfRangeException(nameof(value));
                _position = value;
            }
        }

        public override void Flush() { }

        public override int Read(byte[] buffer, int offset, int count) {
            ThrowIfDisposed();
            int read = _part.Read(_position, buffer, offset, count);
            _position += read;
            return read;
        }

        public override long Seek(long offset, SeekOrigin origin) {
            ThrowIfDisposed();
            long target = origin switch {
                SeekOrigin.Begin => offset,
                SeekOrigin.Current => _position + offset,
                SeekOrigin.End => _part.Length + offset,
                _ => throw new ArgumentOutOfRangeException(nameof(origin))
            };
            if (target < 0) throw new IOException("Cannot seek before the beginning of a package part.");
            _position = target;
            return target;
        }

        public override void SetLength(long value) {
            ThrowIfDisposed();
            if (!_writable) throw new NotSupportedException("The package part stream is read-only.");
            if (value < 0) throw new ArgumentOutOfRangeException(nameof(value));
            _part.SetLength(value);
            if (_position > value) _position = value;
        }

        public override void Write(byte[] buffer, int offset, int count) {
            ThrowIfDisposed();
            if (!_writable) throw new NotSupportedException("The package part stream is read-only.");
            if (count == 0) return;
            _part.Write(_position, buffer, offset, count);
            _position += count;
        }

        protected override void Dispose(bool disposing) {
            _disposed = true;
            base.Dispose(disposing);
        }

        private void ThrowIfDisposed() {
            if (_disposed) throw new ObjectDisposedException(nameof(PartStream));
        }
    }
}
