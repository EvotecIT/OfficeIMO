using System.IO.Compression;
using System.Text;
using System.Xml;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    /// <summary>Rewrites a ZIP package with selected entries changed and every other entry copied byte for byte.</summary>
    internal sealed class ExcelPackageEntrySplicer {
        private const uint LocalHeaderSignature = 0x04034B50;
        private const uint CentralHeaderSignature = 0x02014B50;
        private const uint EndOfCentralDirectorySignature = 0x06054B50;
        private const int LocalHeaderSize = 30;
        private const int CentralHeaderSize = 46;
        private const int EndRecordSize = 22;

        private readonly byte[] _package;
        private readonly int _endRecord;
        private readonly int _directoryOffset;
        private readonly Dictionary<string, Entry> _entries = new Dictionary<string, Entry>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, byte[]> _replacements = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _removals = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private ExcelPackageEntrySplicer(byte[] package, int endRecord, int directoryOffset) {
            _package = package;
            _endRecord = endRecord;
            _directoryOffset = directoryOffset;
        }

        internal IEnumerable<string> Names => _entries.Keys;

        /// <summary>
        /// Returns the package after <paramref name="edit"/>, or null when it is outside the plain ZIP shape handled
        /// here, when <paramref name="edit"/> rejects it with <see cref="InvalidDataException"/>, or when nothing changed.
        /// </summary>
        internal static byte[]? TryEdit(byte[] package, Action<ExcelPackageEntrySplicer> edit) {
            try {
                ExcelPackageEntrySplicer? splicer = Open(package);
                if (splicer == null) return null;
                edit(splicer);
                return splicer.Build();
            } catch (Exception exception) when (exception is InvalidDataException or XmlException or ArgumentException or InvalidOperationException or IndexOutOfRangeException) {
                return null;
            }
        }

        internal bool Contains(string name) => _entries.ContainsKey(name);

        internal bool Matches(string name, byte[] content, int length) {
            Entry entry = _entries[name];
            return ReadUInt32(_package, entry.CentralOffset + 24) == length
                && ReadUInt32(_package, entry.CentralOffset + 16) == Crc(content, length);
        }

        internal void Replace(string name, byte[] content) {
            if (!_entries.ContainsKey(name)) throw new InvalidDataException($"Package entry '{name}' is missing.");
            _replacements[name] = content;
        }

        internal void Remove(string name) {
            if (!_entries.ContainsKey(name)) throw new InvalidDataException($"Package entry '{name}' is missing.");
            _removals.Add(name);
        }

        internal byte[] Read(string name) {
            if (!_entries.TryGetValue(name, out Entry? entry)) throw new InvalidDataException($"Package entry '{name}' is missing.");
            if (ReadUInt32(_package, entry.LocalOffset) != LocalHeaderSignature) throw new InvalidDataException("Invalid local header.");
            int data = entry.LocalOffset + LocalHeaderSize + ReadUInt16(_package, entry.LocalOffset + 26) + ReadUInt16(_package, entry.LocalOffset + 28);
            int method = ReadUInt16(_package, entry.CentralOffset + 10);
            int compressedSize = checked((int)ReadUInt32(_package, entry.CentralOffset + 20));
            var content = new byte[checked((int)ReadUInt32(_package, entry.CentralOffset + 24))];
            if (method == 0) {
                Buffer.BlockCopy(_package, data, content, 0, content.Length);
                return content;
            }
            if (method != 8) throw new InvalidDataException("Unsupported compression.");
            using var compressed = new MemoryStream(_package, data, compressedSize, writable: false);
            using var deflate = new DeflateStream(compressed, CompressionMode.Decompress);
            int read = 0;
            while (read < content.Length) {
                int block = deflate.Read(content, read, content.Length - read);
                if (block == 0) throw new InvalidDataException("Truncated package entry.");
                read += block;
            }
            return content;
        }

        private static ExcelPackageEntrySplicer? Open(byte[] package) {
            int end = FindEndRecord(package);
            if (end < 0) return null;
            int count = ReadUInt16(package, end + 10);
            long directoryOffset = ReadUInt32(package, end + 16);
            if (count == 0xFFFF || directoryOffset == 0xFFFFFFFFL || ReadUInt16(package, end + 4) != 0) return null;

            var splicer = new ExcelPackageEntrySplicer(package, end, (int)directoryOffset);
            int cursor = (int)directoryOffset;
            for (int index = 0; index < count; index++) {
                if (ReadUInt32(package, cursor) != CentralHeaderSignature) return null;
                int nameLength = ReadUInt16(package, cursor + 28);
                int headerLength = CentralHeaderSize + nameLength + ReadUInt16(package, cursor + 30) + ReadUInt16(package, cursor + 32);
                long localOffset = ReadUInt32(package, cursor + 42);
                if (localOffset == 0xFFFFFFFFL) return null;
                string name = Encoding.UTF8.GetString(package, cursor + CentralHeaderSize, nameLength);
                splicer._entries.Add(name, new Entry(cursor, headerLength, (int)localOffset));
                cursor += headerLength;
            }
            return splicer;
        }

        private byte[]? Build() {
            if (_replacements.Count == 0 && _removals.Count == 0) return null;
            Entry[] ordered = _entries.Values.OrderBy(static entry => entry.LocalOffset).ToArray();
            for (int index = 0; index < ordered.Length; index++) {
                ordered[index].LocalEnd = index + 1 < ordered.Length ? ordered[index + 1].LocalOffset : _directoryOffset;
            }

            using var output = new MemoryStream(_package.Length + 512);
            using var directory = new MemoryStream(_endRecord - _directoryOffset + 64);
            int kept = 0;
            foreach (KeyValuePair<string, Entry> pair in _entries.OrderBy(static item => item.Value.LocalOffset)) {
                if (_removals.Contains(pair.Key)) continue;
                kept++;
                Entry entry = pair.Value;
                int localOffset = (int)output.Position;
                if (_replacements.TryGetValue(pair.Key, out byte[]? content)) {
                    WriteReplacement(entry, content, output, directory, localOffset);
                } else {
                    output.Write(_package, entry.LocalOffset, entry.LocalEnd - entry.LocalOffset);
                    var central = new byte[entry.CentralLength];
                    Buffer.BlockCopy(_package, entry.CentralOffset, central, 0, central.Length);
                    WriteUInt32(central, 42, (uint)localOffset);
                    directory.Write(central, 0, central.Length);
                }
            }

            int directoryStart = (int)output.Position;
            directory.Position = 0;
            directory.CopyTo(output);
            int endLength = _package.Length - _endRecord;
            var end = new byte[endLength];
            Buffer.BlockCopy(_package, _endRecord, end, 0, endLength);
            WriteUInt16(end, 8, (ushort)kept);
            WriteUInt16(end, 10, (ushort)kept);
            WriteUInt32(end, 12, (uint)(output.Position - directoryStart));
            WriteUInt32(end, 16, (uint)directoryStart);
            output.Write(end, 0, end.Length);
            return output.ToArray();
        }

        private void WriteReplacement(Entry entry, byte[] content, MemoryStream output, MemoryStream directory, int localOffset) {
            byte[] compressed;
            using (var deflated = new MemoryStream(content.Length / 2 + 64)) {
                using (var deflate = new DeflateStream(deflated, CompressionLevel.Fastest, leaveOpen: true)) {
                    deflate.Write(content, 0, content.Length);
                }
                compressed = deflated.ToArray();
            }

            int nameLength = ReadUInt16(_package, entry.CentralOffset + 28);
            uint crc = (uint)Crc(content, content.Length);
            var local = new byte[LocalHeaderSize];
            WriteUInt32(local, 0, LocalHeaderSignature);
            Buffer.BlockCopy(_package, entry.CentralOffset + 6, local, 4, 2);
            WriteUInt16(local, 6, (ushort)(ReadUInt16(_package, entry.CentralOffset + 8) & ~0x0008));
            WriteUInt16(local, 8, 8);
            Buffer.BlockCopy(_package, entry.CentralOffset + 12, local, 10, 4);
            WriteUInt32(local, 14, crc);
            WriteUInt32(local, 18, (uint)compressed.Length);
            WriteUInt32(local, 22, (uint)content.Length);
            WriteUInt16(local, 26, (ushort)nameLength);
            output.Write(local, 0, local.Length);
            output.Write(_package, entry.CentralOffset + CentralHeaderSize, nameLength);
            output.Write(compressed, 0, compressed.Length);

            var central = new byte[entry.CentralLength];
            Buffer.BlockCopy(_package, entry.CentralOffset, central, 0, central.Length);
            WriteUInt16(central, 8, (ushort)(ReadUInt16(central, 8) & ~0x0008));
            WriteUInt16(central, 10, 8);
            WriteUInt32(central, 16, crc);
            WriteUInt32(central, 20, (uint)compressed.Length);
            WriteUInt32(central, 24, (uint)content.Length);
            WriteUInt32(central, 42, (uint)localOffset);
            directory.Write(central, 0, central.Length);
        }

        private static long Crc(byte[] content, int length) =>
            OfficePngCrc32.Complete(OfficePngCrc32.Append(OfficePngCrc32.Begin(), content, 0, length));

        private static int FindEndRecord(byte[] package) {
            int lowest = Math.Max(0, package.Length - EndRecordSize - ushort.MaxValue);
            for (int index = package.Length - EndRecordSize; index >= lowest; index--) {
                if (ReadUInt32(package, index) == EndOfCentralDirectorySignature
                    && index + EndRecordSize + ReadUInt16(package, index + 20) == package.Length) {
                    return index;
                }
            }
            return -1;
        }

        private static int ReadUInt16(byte[] data, int offset) => data[offset] | data[offset + 1] << 8;

        private static long ReadUInt32(byte[] data, int offset) =>
            (uint)(data[offset] | data[offset + 1] << 8 | data[offset + 2] << 16 | data[offset + 3] << 24);

        private static void WriteUInt16(byte[] data, int offset, ushort value) {
            data[offset] = (byte)value;
            data[offset + 1] = (byte)(value >> 8);
        }

        private static void WriteUInt32(byte[] data, int offset, uint value) {
            WriteUInt16(data, offset, (ushort)value);
            WriteUInt16(data, offset + 2, (ushort)(value >> 16));
        }

        private sealed class Entry {
            internal Entry(int centralOffset, int centralLength, int localOffset) {
                CentralOffset = centralOffset;
                CentralLength = centralLength;
                LocalOffset = localOffset;
            }

            internal int CentralOffset { get; }
            internal int CentralLength { get; }
            internal int LocalOffset { get; }
            internal int LocalEnd { get; set; }
        }
    }
}
