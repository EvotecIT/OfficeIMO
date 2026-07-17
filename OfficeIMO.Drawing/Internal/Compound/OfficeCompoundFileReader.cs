using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace OfficeIMO.Drawing.Internal {
    /// <summary>
    /// Reads OLE compound document containers used by legacy Office binary formats.
    /// </summary>
    internal static partial class OfficeCompoundFileReader {
        private const int HeaderSize = 512;
        private const int MiniSectorSize = 64;
        private const int DirectoryEntrySize = 128;
        private const uint FreeSect = 0xffffffff;
        private const uint EndOfChain = 0xfffffffe;
        private static readonly byte[] Signature = { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 };

        internal static bool TryRead(byte[] bytes, out Dictionary<string, byte[]> streams, out string? error) {
            bool result = TryRead(bytes, OfficeCompoundReadOptions.Default, out OfficeCompoundFile? compoundFile, out error);
            streams = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            if (result && compoundFile != null) {
                foreach (KeyValuePair<string, byte[]> stream in compoundFile.Streams) {
                    streams[stream.Key] = stream.Value;
                }
            }

            return result;
        }

        internal static bool TryRead(byte[] bytes, out OfficeCompoundFile? compoundFile, out string? error) {
            return TryRead(bytes, OfficeCompoundReadOptions.Default, out compoundFile, out error);
        }

        internal static bool TryRead(byte[] bytes, OfficeCompoundReadOptions options,
            out OfficeCompoundFile? compoundFile, out string? error) {
            compoundFile = null;
            error = null;

            try {
                if (bytes == null || bytes.Length < HeaderSize || !HasSignature(bytes)) {
                    error = "The file does not start with the OLE compound document signature.";
                    return false;
                }
                if (options == null) throw new ArgumentNullException(nameof(options));

                var streams = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
                ushort majorVersion = ReadUInt16(bytes, 26);
                ushort byteOrder = ReadUInt16(bytes, 28);
                ushort sectorShift = ReadUInt16(bytes, 30);
                ushort miniSectorShift = ReadUInt16(bytes, 32);
                if ((sectorShift != 9 && sectorShift != 12) || miniSectorShift != 6) {
                    error = $"Unsupported compound file sector sizes. SectorShift={sectorShift}, MiniSectorShift={miniSectorShift}.";
                    return false;
                }
                int sectorSize = 1 << sectorShift;
                int miniSectorSize = 1 << miniSectorShift;
                bool validVersion = (majorVersion == 3 && sectorSize == 512) || (majorVersion == 4 && sectorSize == 4096);
                if (!validVersion || byteOrder != 0xfffe) {
                    error = $"The OLE compound file could not be read. The header has an unsupported version or byte order. Version={majorVersion}, ByteOrder=0x{byteOrder:x4}.";
                    return false;
                }
                if (miniSectorSize != MiniSectorSize || bytes.Length < sectorSize) {
                    error = "The compound file is shorter than its declared header sector.";
                    return false;
                }

                int fatSectorCount = checked((int)ReadUInt32(bytes, 44));
                uint directoryStart = ReadUInt32(bytes, 48);
                uint miniCutoff = ReadUInt32(bytes, 56);
                uint miniFatStart = ReadUInt32(bytes, 60);
                int miniFatSectorCount = checked((int)ReadUInt32(bytes, 64));
                uint firstDifat = ReadUInt32(bytes, 68);
                int difatSectorCount = checked((int)ReadUInt32(bytes, 72));

                int physicalSectorCount = (bytes.Length - sectorSize) / sectorSize;
                if (fatSectorCount < 0 || fatSectorCount > physicalSectorCount || difatSectorCount < 0 ||
                    difatSectorCount > physicalSectorCount) {
                    throw new InvalidDataException("Compound file allocation table counts exceed the file size.");
                }

                List<uint> fatSectorIds = ReadDifat(bytes, sectorSize, firstDifat, difatSectorCount, fatSectorCount);
                uint[] fat = ReadFat(bytes, sectorSize, fatSectorIds);
                long maximumDirectoryBytes = checked((long)options.MaxDirectoryEntries * DirectoryEntrySize);
                byte[] directoryBytes = ReadRegularStream(bytes, sectorSize, fat, directoryStart, long.MaxValue,
                    maximumDirectoryBytes);
                List<DirectoryEntry> entries = ReadDirectoryEntries(directoryBytes, majorVersion, options.MaxDirectoryEntries);
                DirectoryEntry? root = entries.FirstOrDefault(entry => entry.ObjectType == 5);
                if (root == null) throw new InvalidDataException("Compound file root directory entry is missing.");
                if (root.Size < 0 || root.Size > options.MaxTotalStreamBytes || root.Size > int.MaxValue) {
                    throw new InvalidDataException("Compound file mini stream exceeds configured bounds.");
                }

                IReadOnlyDictionary<int, string> streamPaths = BuildCompoundEntryPaths(entries);
                DirectoryEntry[] streamEntries = entries.Where(entry => entry.ObjectType == 2).ToArray();
                if (streamEntries.Length > options.MaxStreamCount) {
                    throw new InvalidDataException($"Compound file stream count {streamEntries.Length} exceeds {options.MaxStreamCount}.");
                }
                long totalStreamBytes = 0;
                foreach (DirectoryEntry entry in streamEntries) {
                    string path = streamPaths.TryGetValue(entry.Index, out string? entryPath)
                        ? entryPath
                        : entry.Name;
                    if (entry.Size < 0 || entry.Size > options.MaxStreamBytes || entry.Size > int.MaxValue) {
                        throw new InvalidDataException($"Compound stream '{entry.Name}' has unsupported size {entry.Size}.");
                    }
                    totalStreamBytes = checked(totalStreamBytes + entry.Size);
                    if (totalStreamBytes > options.MaxTotalStreamBytes) {
                        throw new InvalidDataException($"Compound stream bytes exceed {options.MaxTotalStreamBytes}.");
                    }
                    options.StreamSizeValidator?.Invoke(path, entry.Size);
                }

                byte[] miniStream = root == null || root.StartSector == EndOfChain
                    ? Array.Empty<byte>()
                    : ReadRegularStream(bytes, sectorSize, fat, root.StartSector, root.Size);
                uint[] miniFat = miniFatStart == EndOfChain || miniFatSectorCount == 0
                    ? Array.Empty<uint>()
                    : BytesToUInt32Array(ReadRegularStream(bytes, sectorSize, fat, miniFatStart, (long)miniFatSectorCount * sectorSize));

                foreach (DirectoryEntry entry in streamEntries) {
                    string path = streamPaths.TryGetValue(entry.Index, out string? entryPath)
                        ? entryPath
                        : entry.Name;
                    byte[] data = entry.Size < miniCutoff
                        ? ReadMiniStream(miniStream, miniFat, entry.StartSector, entry.Size)
                        : ReadRegularStream(bytes, sectorSize, fat, entry.StartSector, entry.Size);
                    streams[path] = data;
                    if (string.Equals(path, entry.Name, StringComparison.OrdinalIgnoreCase)) {
                        streams[entry.Name] = data;
                    }
                }

                compoundFile = new OfficeCompoundFile(streams, BuildCompoundEntries(entries),
                    CreateCompoundEntry(root!, "Root Entry"));
                return true;
            } catch (Exception ex) when (ex is IOException || ex is ArgumentException || ex is InvalidDataException || ex is OverflowException || ex is IndexOutOfRangeException) {
                compoundFile = null;
                error = $"The OLE compound file could not be read. {ex.Message}";
                return false;
            }
        }

        private static bool HasSignature(byte[] bytes) {
            for (int i = 0; i < Signature.Length; i++) {
                if (bytes[i] != Signature[i]) return false;
            }

            return true;
        }

        private static List<uint> ReadDifat(byte[] bytes, int sectorSize, uint firstDifat, int difatSectorCount, int fatSectorCount) {
            var result = new List<uint>(fatSectorCount);
            var visitedFatSectors = new HashSet<uint>();
            for (int i = 0; i < 109 && result.Count < fatSectorCount; i++) {
                uint sector = ReadUInt32(bytes, 76 + i * 4);
                if (sector != FreeSect && !visitedFatSectors.Add(sector)) throw new InvalidDataException("Duplicate FAT sector reference.");
                if (sector != FreeSect) result.Add(sector);
            }

            uint next = firstDifat;
            var visitedDifatSectors = new HashSet<uint>();
            int entriesPerSector = sectorSize / 4 - 1;
            for (int d = 0; d < difatSectorCount && next != EndOfChain && result.Count < fatSectorCount; d++) {
                if (!visitedDifatSectors.Add(next)) throw new InvalidDataException("Cyclic DIFAT sector chain.");
                int offset = SectorOffset(next, sectorSize);
                for (int i = 0; i < entriesPerSector && result.Count < fatSectorCount; i++) {
                    uint sector = ReadUInt32(bytes, offset + i * 4);
                    if (sector != FreeSect && !visitedFatSectors.Add(sector)) throw new InvalidDataException("Duplicate FAT sector reference.");
                    if (sector != FreeSect) result.Add(sector);
                }

                next = ReadUInt32(bytes, offset + entriesPerSector * 4);
            }

            if (result.Count != fatSectorCount) throw new InvalidDataException("The DIFAT does not reference the declared FAT sector count.");

            return result;
        }

        private static uint[] ReadFat(byte[] bytes, int sectorSize, List<uint> fatSectorIds) {
            var entries = new List<uint>(fatSectorIds.Count * (sectorSize / 4));
            foreach (uint sector in fatSectorIds) {
                int offset = SectorOffset(sector, sectorSize);
                for (int i = 0; i < sectorSize / 4; i++) {
                    entries.Add(ReadUInt32(bytes, offset + i * 4));
                }
            }

            return entries.ToArray();
        }

        private static byte[] ReadRegularStream(byte[] bytes, int sectorSize, uint[] fat, uint startSector, long size,
            long maximumBytes = long.MaxValue) {
            if (size == 0) return Array.Empty<byte>();
            if (startSector == EndOfChain) {
                if (size == long.MaxValue) return Array.Empty<byte>();
                throw new InvalidDataException("A non-empty compound stream has no sector chain.");
            }

            using var output = new MemoryStream();
            uint sector = startSector;
            var visited = new HashSet<uint>();
            while (sector != EndOfChain && sector != FreeSect) {
                if (sector >= fat.Length || !visited.Add(sector)) {
                    throw new InvalidDataException("Invalid compound file sector chain.");
                }

                int offset = SectorOffset(sector, sectorSize);
                if (offset < 0 || offset + sectorSize > bytes.Length) {
                    throw new InvalidDataException("Compound file sector points outside the file.");
                }
                if (output.Length > maximumBytes - sectorSize) {
                    throw new InvalidDataException($"Compound directory entry count exceeds {maximumBytes / DirectoryEntrySize}.");
                }

                output.Write(bytes, offset, sectorSize);
                sector = fat[sector];
                if (size != long.MaxValue && output.Length >= size) {
                    break;
                }
            }

            byte[] data = output.ToArray();
            if (size != long.MaxValue && data.LongLength < size) {
                throw new InvalidDataException("Compound file sector chain is shorter than the declared stream size.");
            }
            if (size != long.MaxValue && data.LongLength > size) {
                Array.Resize(ref data, checked((int)size));
            }

            return data;
        }

        private static byte[] ReadMiniStream(byte[] miniStream, uint[] miniFat, uint startSector, long size) {
            if (size == 0) return Array.Empty<byte>();
            if (startSector == EndOfChain) throw new InvalidDataException("A non-empty mini stream has no sector chain.");

            using var output = new MemoryStream();
            uint sector = startSector;
            var visited = new HashSet<uint>();
            while (sector != EndOfChain && sector != FreeSect) {
                if (sector >= miniFat.Length || !visited.Add(sector)) {
                    throw new InvalidDataException("Invalid compound file mini sector chain.");
                }

                int offset = checked((int)sector * MiniSectorSize);
                if (offset < 0 || offset > miniStream.Length) {
                    throw new InvalidDataException("Compound file mini sector points outside the mini stream.");
                }

                output.Write(miniStream, offset, Math.Min(MiniSectorSize, miniStream.Length - offset));
                sector = miniFat[sector];
                if (output.Length >= size) {
                    break;
                }
            }

            byte[] data = output.ToArray();
            if (data.LongLength < size) throw new InvalidDataException("Mini sector chain is shorter than the declared stream size.");
            if (data.LongLength > size) {
                Array.Resize(ref data, checked((int)size));
            }

            return data;
        }

        private static List<DirectoryEntry> ReadDirectoryEntries(
            byte[] directoryBytes, ushort majorVersion, int maximumEntries,
            CancellationToken cancellationToken = default) {
            var result = new List<DirectoryEntry>();
            for (int offset = 0; offset + DirectoryEntrySize <= directoryBytes.Length; offset += DirectoryEntrySize) {
                cancellationToken.ThrowIfCancellationRequested();
                if (result.Count >= maximumEntries) throw new InvalidDataException($"Compound directory entry count exceeds {maximumEntries}.");
                ushort nameLength = ReadUInt16(directoryBytes, offset + 64);
                byte objectType = directoryBytes[offset + 66];
                string name = objectType == 0 || nameLength < 2 || nameLength > 64
                    ? string.Empty
                    : Encoding.Unicode.GetString(directoryBytes, offset, nameLength - 2);

                result.Add(new DirectoryEntry(
                    result.Count,
                    name,
                    objectType,
                    ReadUInt32(directoryBytes, offset + 68),
                    ReadUInt32(directoryBytes, offset + 72),
                    ReadUInt32(directoryBytes, offset + 76),
                    ReadGuid(directoryBytes, offset + 80),
                    ReadUInt32(directoryBytes, offset + 96),
                    ReadUInt64(directoryBytes, offset + 100),
                    ReadUInt64(directoryBytes, offset + 108),
                    ReadUInt32(directoryBytes, offset + 116),
                    majorVersion == 3
                        ? ReadUInt32(directoryBytes, offset + 120)
                        : checked((long)ReadUInt64(directoryBytes, offset + 120))));
            }

            return result;
        }

        private static IReadOnlyList<OfficeCompoundFileEntry>
            BuildCompoundEntries(IReadOnlyList<DirectoryEntry> entries,
                CancellationToken cancellationToken = default) {
            var result = new List<OfficeCompoundFileEntry>();
            DirectoryEntry? root = entries.FirstOrDefault(entry => entry.ObjectType == 5);
            if (root != null) {
                TraverseDirectoryTree(entries, root.ChildId, string.Empty,
                    result, new HashSet<int>(), 0, cancellationToken);
            }

            var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (OfficeCompoundFileEntry entry in result) {
                cancellationToken.ThrowIfCancellationRequested();
                paths.Add(entry.Path);
            }
            foreach (DirectoryEntry entry in entries) {
                cancellationToken.ThrowIfCancellationRequested();
                if (entry.ObjectType == 0 || string.IsNullOrEmpty(entry.Name)) {
                    continue;
                }

                if (paths.Add(entry.Name)) {
                    result.Add(CreateCompoundEntry(entry, entry.Name, isFallback: true));
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            return result
                .OrderBy(entry => entry.Path, StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }

        private static IReadOnlyDictionary<int, string> BuildCompoundEntryPaths(IReadOnlyList<DirectoryEntry> entries) {
            var result = new Dictionary<int, string>();
            DirectoryEntry? root = entries.FirstOrDefault(entry => entry.ObjectType == 5);
            if (root != null) {
                TraverseDirectoryTree(entries, root.ChildId, string.Empty, result, new HashSet<int>(), 0);
            }

            foreach (DirectoryEntry entry in entries) {
                if (entry.ObjectType == 0 || string.IsNullOrEmpty(entry.Name)) {
                    continue;
                }

                if (!result.ContainsKey(entry.Index)) {
                    result[entry.Index] = entry.Name;
                }
            }

            return result;
        }

        private static void TraverseDirectoryTree(
            IReadOnlyList<DirectoryEntry> entries,
            uint entryId,
            string parentPath,
            List<OfficeCompoundFileEntry> result,
            HashSet<int> visited,
            int depth,
            CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            if (depth > 256) throw new InvalidDataException("Compound directory tree exceeds the supported depth.");
            if (entryId == FreeSect || entryId == EndOfChain || entryId >= entries.Count) {
                return;
            }

            DirectoryEntry entry = entries[(int)entryId];
            if (!visited.Add(entry.Index)) {
                return;
            }

            TraverseDirectoryTree(entries, entry.LeftSiblingId, parentPath,
                result, visited, depth + 1, cancellationToken);
            if (entry.ObjectType != 0 && !string.IsNullOrEmpty(entry.Name)) {
                string path = string.IsNullOrEmpty(parentPath) ? entry.Name : parentPath + "/" + entry.Name;
                result.Add(CreateCompoundEntry(entry, path));
                if (entry.ObjectType == 1 || entry.ObjectType == 5) {
                    TraverseDirectoryTree(entries, entry.ChildId, path,
                        result, visited, depth + 1, cancellationToken);
                }
            }

            TraverseDirectoryTree(entries, entry.RightSiblingId, parentPath,
                result, visited, depth + 1, cancellationToken);
        }

        private static void TraverseDirectoryTree(
            IReadOnlyList<DirectoryEntry> entries,
            uint entryId,
            string parentPath,
            Dictionary<int, string> result,
            HashSet<int> visited,
            int depth) {
            if (depth > 256) throw new InvalidDataException("Compound directory tree exceeds the supported depth.");
            if (entryId == FreeSect || entryId == EndOfChain || entryId >= entries.Count) {
                return;
            }

            DirectoryEntry entry = entries[(int)entryId];
            if (!visited.Add(entry.Index)) {
                return;
            }

            TraverseDirectoryTree(entries, entry.LeftSiblingId, parentPath, result, visited, depth + 1);
            if (entry.ObjectType != 0 && !string.IsNullOrEmpty(entry.Name)) {
                string path = string.IsNullOrEmpty(parentPath) ? entry.Name : parentPath + "/" + entry.Name;
                result[entry.Index] = path;
                if (entry.ObjectType == 1 || entry.ObjectType == 5) {
                    TraverseDirectoryTree(entries, entry.ChildId, path, result, visited, depth + 1);
                }
            }

            TraverseDirectoryTree(entries, entry.RightSiblingId, parentPath, result, visited, depth + 1);
        }

        private static uint[] BytesToUInt32Array(byte[] bytes) {
            uint[] result = new uint[bytes.Length / 4];
            for (int i = 0; i < result.Length; i++) {
                result[i] = ReadUInt32(bytes, i * 4);
            }

            return result;
        }

        private static int SectorOffset(uint sector, int sectorSize) {
            // CFB version 4 pads the fixed 512-byte header structure to one full 4096-byte header sector.
            // In every CFB version, sector 0 therefore starts at one declared sector size from the file start.
            return checked(((int)sector + 1) * sectorSize);
        }

        private static ushort ReadUInt16(byte[] bytes, int offset) {
            if (offset < 0 || offset + 2 > bytes.Length) throw new InvalidDataException("Unexpected end of compound file.");
            return (ushort)(bytes[offset] | (bytes[offset + 1] << 8));
        }

        private static uint ReadUInt32(byte[] bytes, int offset) {
            if (offset < 0 || offset + 4 > bytes.Length) throw new InvalidDataException("Unexpected end of compound file.");
            return (uint)(bytes[offset]
                | (bytes[offset + 1] << 8)
                | (bytes[offset + 2] << 16)
                | (bytes[offset + 3] << 24));
        }

        private static ulong ReadUInt64(byte[] bytes, int offset) {
            uint low = ReadUInt32(bytes, offset);
            uint high = ReadUInt32(bytes, offset + 4);
            return low | ((ulong)high << 32);
        }

        private static Guid ReadGuid(byte[] bytes, int offset) {
            var value = new byte[16];
            Buffer.BlockCopy(bytes, offset, value, 0, value.Length);
            return new Guid(value);
        }

        private static OfficeCompoundFileEntry CreateCompoundEntry(DirectoryEntry entry, string path,
            bool isFallback = false) => new OfficeCompoundFileEntry(entry.Name, path, entry.ObjectType, entry.Size,
                isFallback, entry.ClassId, entry.StateBits, entry.CreationTime, entry.ModifiedTime);

        private sealed class DirectoryEntry {
            internal DirectoryEntry(int index, string name, byte objectType, uint leftSiblingId,
                uint rightSiblingId, uint childId, Guid classId, uint stateBits, ulong creationTime,
                ulong modifiedTime, uint startSector, long size) {
                Index = index;
                Name = name;
                ObjectType = objectType;
                LeftSiblingId = leftSiblingId;
                RightSiblingId = rightSiblingId;
                ChildId = childId;
                ClassId = classId;
                StateBits = stateBits;
                CreationTime = creationTime;
                ModifiedTime = modifiedTime;
                StartSector = startSector;
                Size = size;
            }

            internal int Index { get; }

            internal string Name { get; }

            internal byte ObjectType { get; }

            internal uint LeftSiblingId { get; }

            internal uint RightSiblingId { get; }

            internal uint ChildId { get; }

            internal Guid ClassId { get; }

            internal uint StateBits { get; }

            internal ulong CreationTime { get; }

            internal ulong ModifiedTime { get; }

            internal uint StartSector { get; }

            internal long Size { get; }
        }
    }

    /// <summary>
    /// Resource limits used while decoding a compound file.
    /// </summary>
    internal sealed class OfficeCompoundReadOptions {
        internal static OfficeCompoundReadOptions Default { get; } = new OfficeCompoundReadOptions();

        internal OfficeCompoundReadOptions(int maxDirectoryEntries = 65536, int maxStreamCount = 32768,
            long maxStreamBytes = 256L * 1024L * 1024L, long maxTotalStreamBytes = 512L * 1024L * 1024L,
            Action<string, long>? streamSizeValidator = null) {
            if (maxDirectoryEntries <= 0) throw new ArgumentOutOfRangeException(nameof(maxDirectoryEntries));
            if (maxStreamCount <= 0) throw new ArgumentOutOfRangeException(nameof(maxStreamCount));
            if (maxStreamBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxStreamBytes));
            if (maxTotalStreamBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxTotalStreamBytes));
            MaxDirectoryEntries = maxDirectoryEntries;
            MaxStreamCount = maxStreamCount;
            MaxStreamBytes = maxStreamBytes;
            MaxTotalStreamBytes = maxTotalStreamBytes;
            StreamSizeValidator = streamSizeValidator;
        }

        internal int MaxDirectoryEntries { get; }

        internal int MaxStreamCount { get; }

        internal long MaxStreamBytes { get; }

        internal long MaxTotalStreamBytes { get; }

        internal Action<string, long>? StreamSizeValidator { get; }
    }

    /// <summary>Signals a path-aware stream-size rejection before compound stream bytes are buffered.</summary>
    internal sealed class OfficeCompoundStreamLimitExceededException : Exception {
        internal OfficeCompoundStreamLimitExceededException(string limitName, long actualValue, long maximumValue)
            : base($"{limitName} exceeded: {actualValue} is greater than {maximumValue}.") {
            LimitName = limitName;
            ActualValue = actualValue;
            MaximumValue = maximumValue;
        }

        internal string LimitName { get; }

        internal long ActualValue { get; }

        internal long MaximumValue { get; }
    }
}
