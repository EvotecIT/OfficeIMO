using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeIMO.Shared {
    /// <summary>
    /// Reads OLE compound document containers used by legacy Office binary formats.
    /// </summary>
    internal static class OfficeCompoundFileReader {
        private const int HeaderSize = 512;
        private const int MiniSectorSize = 64;
        private const int DirectoryEntrySize = 128;
        private const uint FreeSect = 0xffffffff;
        private const uint EndOfChain = 0xfffffffe;
        private static readonly byte[] Signature = { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 };

        internal static bool TryRead(byte[] bytes, out Dictionary<string, byte[]> streams, out string? error) {
            bool result = TryRead(bytes, out OfficeCompoundFile? compoundFile, out error);
            streams = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            if (result && compoundFile != null) {
                foreach (KeyValuePair<string, byte[]> stream in compoundFile.Streams) {
                    streams[stream.Key] = stream.Value;
                }
            }

            return result;
        }

        internal static bool TryRead(byte[] bytes, out OfficeCompoundFile? compoundFile, out string? error) {
            compoundFile = null;
            error = null;

            try {
                if (bytes.Length < HeaderSize || !HasSignature(bytes)) {
                    error = "The file does not start with the OLE compound document signature.";
                    return false;
                }

                var streams = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
                ushort sectorShift = ReadUInt16(bytes, 30);
                ushort miniSectorShift = ReadUInt16(bytes, 32);
                int sectorSize = 1 << sectorShift;
                int miniSectorSize = 1 << miniSectorShift;
                if (sectorSize < 512 || sectorSize > 4096 || miniSectorSize != MiniSectorSize) {
                    error = $"Unsupported compound file sector sizes. Sector={sectorSize}, MiniSector={miniSectorSize}.";
                    return false;
                }

                int fatSectorCount = checked((int)ReadUInt32(bytes, 44));
                uint directoryStart = ReadUInt32(bytes, 48);
                uint miniCutoff = ReadUInt32(bytes, 56);
                uint miniFatStart = ReadUInt32(bytes, 60);
                int miniFatSectorCount = checked((int)ReadUInt32(bytes, 64));
                uint firstDifat = ReadUInt32(bytes, 68);
                int difatSectorCount = checked((int)ReadUInt32(bytes, 72));

                List<uint> fatSectorIds = ReadDifat(bytes, sectorSize, firstDifat, difatSectorCount, fatSectorCount);
                uint[] fat = ReadFat(bytes, sectorSize, fatSectorIds);
                byte[] directoryBytes = ReadRegularStream(bytes, sectorSize, fat, directoryStart, long.MaxValue);
                List<DirectoryEntry> entries = ReadDirectoryEntries(directoryBytes);
                DirectoryEntry? root = entries.FirstOrDefault(entry => entry.ObjectType == 5);
                byte[] miniStream = root == null || root.StartSector == EndOfChain
                    ? Array.Empty<byte>()
                    : ReadRegularStream(bytes, sectorSize, fat, root.StartSector, root.Size);
                uint[] miniFat = miniFatStart == EndOfChain || miniFatSectorCount == 0
                    ? Array.Empty<uint>()
                    : BytesToUInt32Array(ReadRegularStream(bytes, sectorSize, fat, miniFatStart, (long)miniFatSectorCount * sectorSize));

                IReadOnlyDictionary<int, string> streamPaths = BuildCompoundEntryPaths(entries);
                foreach (DirectoryEntry entry in entries.Where(entry => entry.ObjectType == 2)) {
                    byte[] data = entry.Size < miniCutoff
                        ? ReadMiniStream(miniStream, miniFat, entry.StartSector, entry.Size)
                        : ReadRegularStream(bytes, sectorSize, fat, entry.StartSector, entry.Size);
                    string path = streamPaths.TryGetValue(entry.Index, out string? entryPath)
                        ? entryPath
                        : entry.Name;
                    streams[path] = data;
                    if (string.Equals(path, entry.Name, StringComparison.OrdinalIgnoreCase)) {
                        streams[entry.Name] = data;
                    }
                }

                compoundFile = new OfficeCompoundFile(streams, BuildCompoundEntries(entries));
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
            for (int i = 0; i < 109 && result.Count < fatSectorCount; i++) {
                uint sector = ReadUInt32(bytes, 76 + i * 4);
                if (sector != FreeSect) result.Add(sector);
            }

            uint next = firstDifat;
            for (int d = 0; d < difatSectorCount && next != EndOfChain && result.Count < fatSectorCount; d++) {
                int offset = SectorOffset(next, sectorSize);
                for (int i = 0; i < 127 && result.Count < fatSectorCount; i++) {
                    uint sector = ReadUInt32(bytes, offset + i * 4);
                    if (sector != FreeSect) result.Add(sector);
                }

                next = ReadUInt32(bytes, offset + 127 * 4);
            }

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

        private static byte[] ReadRegularStream(byte[] bytes, int sectorSize, uint[] fat, uint startSector, long size) {
            if (startSector == EndOfChain) return Array.Empty<byte>();

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

                output.Write(bytes, offset, sectorSize);
                sector = fat[sector];
                if (size != long.MaxValue && output.Length >= size) {
                    break;
                }
            }

            byte[] data = output.ToArray();
            if (size != long.MaxValue && data.LongLength > size) {
                Array.Resize(ref data, checked((int)size));
            }

            return data;
        }

        private static byte[] ReadMiniStream(byte[] miniStream, uint[] miniFat, uint startSector, long size) {
            if (startSector == EndOfChain) return Array.Empty<byte>();

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
            if (data.LongLength > size) {
                Array.Resize(ref data, checked((int)size));
            }

            return data;
        }

        private static List<DirectoryEntry> ReadDirectoryEntries(byte[] directoryBytes) {
            var result = new List<DirectoryEntry>();
            for (int offset = 0; offset + DirectoryEntrySize <= directoryBytes.Length; offset += DirectoryEntrySize) {
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
                    ReadUInt32(directoryBytes, offset + 116),
                    unchecked((long)ReadUInt64(directoryBytes, offset + 120))));
            }

            return result;
        }

        private static IReadOnlyList<OfficeCompoundFileEntry> BuildCompoundEntries(IReadOnlyList<DirectoryEntry> entries) {
            var result = new List<OfficeCompoundFileEntry>();
            DirectoryEntry? root = entries.FirstOrDefault(entry => entry.ObjectType == 5);
            if (root != null) {
                TraverseDirectoryTree(entries, root.ChildId, string.Empty, result, new HashSet<int>());
            }

            foreach (DirectoryEntry entry in entries) {
                if (entry.ObjectType == 0 || string.IsNullOrEmpty(entry.Name)) {
                    continue;
                }

                if (!result.Any(item => string.Equals(item.Path, entry.Name, StringComparison.OrdinalIgnoreCase))) {
                    result.Add(new OfficeCompoundFileEntry(entry.Name, entry.Name, entry.ObjectType, entry.Size));
                }
            }

            return result
                .OrderBy(entry => entry.Path, StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }

        private static IReadOnlyDictionary<int, string> BuildCompoundEntryPaths(IReadOnlyList<DirectoryEntry> entries) {
            var result = new Dictionary<int, string>();
            DirectoryEntry? root = entries.FirstOrDefault(entry => entry.ObjectType == 5);
            if (root != null) {
                TraverseDirectoryTree(entries, root.ChildId, string.Empty, result, new HashSet<int>());
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
            HashSet<int> visited) {
            if (entryId == FreeSect || entryId == EndOfChain || entryId >= entries.Count) {
                return;
            }

            DirectoryEntry entry = entries[(int)entryId];
            if (!visited.Add(entry.Index)) {
                return;
            }

            TraverseDirectoryTree(entries, entry.LeftSiblingId, parentPath, result, visited);
            if (entry.ObjectType != 0 && !string.IsNullOrEmpty(entry.Name)) {
                string path = string.IsNullOrEmpty(parentPath) ? entry.Name : parentPath + "/" + entry.Name;
                result.Add(new OfficeCompoundFileEntry(entry.Name, path, entry.ObjectType, entry.Size));
                if (entry.ObjectType == 1 || entry.ObjectType == 5) {
                    TraverseDirectoryTree(entries, entry.ChildId, path, result, visited);
                }
            }

            TraverseDirectoryTree(entries, entry.RightSiblingId, parentPath, result, visited);
        }

        private static void TraverseDirectoryTree(
            IReadOnlyList<DirectoryEntry> entries,
            uint entryId,
            string parentPath,
            Dictionary<int, string> result,
            HashSet<int> visited) {
            if (entryId == FreeSect || entryId == EndOfChain || entryId >= entries.Count) {
                return;
            }

            DirectoryEntry entry = entries[(int)entryId];
            if (!visited.Add(entry.Index)) {
                return;
            }

            TraverseDirectoryTree(entries, entry.LeftSiblingId, parentPath, result, visited);
            if (entry.ObjectType != 0 && !string.IsNullOrEmpty(entry.Name)) {
                string path = string.IsNullOrEmpty(parentPath) ? entry.Name : parentPath + "/" + entry.Name;
                result[entry.Index] = path;
                if (entry.ObjectType == 1 || entry.ObjectType == 5) {
                    TraverseDirectoryTree(entries, entry.ChildId, path, result, visited);
                }
            }

            TraverseDirectoryTree(entries, entry.RightSiblingId, parentPath, result, visited);
        }

        private static uint[] BytesToUInt32Array(byte[] bytes) {
            uint[] result = new uint[bytes.Length / 4];
            for (int i = 0; i < result.Length; i++) {
                result[i] = ReadUInt32(bytes, i * 4);
            }

            return result;
        }

        private static int SectorOffset(uint sector, int sectorSize) {
            return checked(HeaderSize + (int)sector * sectorSize);
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

        private sealed class DirectoryEntry {
            internal DirectoryEntry(int index, string name, byte objectType, uint leftSiblingId, uint rightSiblingId, uint childId, uint startSector, long size) {
                Index = index;
                Name = name;
                ObjectType = objectType;
                LeftSiblingId = leftSiblingId;
                RightSiblingId = rightSiblingId;
                ChildId = childId;
                StartSector = startSector;
                Size = size;
            }

            internal int Index { get; }

            internal string Name { get; }

            internal byte ObjectType { get; }

            internal uint LeftSiblingId { get; }

            internal uint RightSiblingId { get; }

            internal uint ChildId { get; }

            internal uint StartSector { get; }

            internal long Size { get; }
        }
    }
}
