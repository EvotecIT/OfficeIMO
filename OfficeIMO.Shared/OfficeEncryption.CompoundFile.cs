#nullable enable
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Shared {
    internal static partial class OfficeEncryption {
        private sealed class CompoundFile {
            private const int SectorSize = 512;
            private const int MiniSectorSize = 64;
            private const int DirectoryEntrySize = 128;
            private const int MiniStreamCutoff = 4096;
            private const uint FreeSect = 0xffffffff;
            private const uint EndOfChain = 0xfffffffe;
            private const uint FatSect = 0xfffffffd;
            private const uint DifSect = 0xfffffffc;

            public static byte[] Write(Dictionary<string, byte[]> streams) {
                var streamInfos = streams
                    .OrderBy(kvp => kvp.Key, StringComparer.Ordinal)
                    .Select(kvp => new StreamInfo(kvp.Key, kvp.Value))
                    .ToList();

                var miniStreams = streamInfos.Where(s => s.Data.Length < MiniStreamCutoff).ToList();
                var regularStreams = streamInfos.Where(s => s.Data.Length >= MiniStreamCutoff).ToList();

                byte[] miniStreamBytes = BuildMiniStream(miniStreams);
                byte[] miniFatBytes = BuildMiniFat(miniStreams);

                var regularChains = new List<RegularChain>();
                foreach (var info in regularStreams) {
                    regularChains.Add(new RegularChain(info, SplitIntoSectors(info.Data)));
                }

                var miniStreamSectors = SplitIntoSectors(miniStreamBytes);
                var miniFatSectors = SplitIntoSectors(miniFatBytes);
                int directorySectorCount = SplitIntoSectors(BuildDirectory(streamInfos, miniStreamBytes.Length, 0)).Count;

                int dataSectorCount = regularChains.Sum(c => c.Sectors.Count) + miniStreamSectors.Count + miniFatSectors.Count + directorySectorCount;
                int fatSectorCount = 0;
                int difatSectorCount = 0;
                while (true) {
                    int nextFatCount = CeilingDiv(dataSectorCount + fatSectorCount + difatSectorCount, SectorSize / 4);
                    int nextDifatCount = nextFatCount <= 109 ? 0 : CeilingDiv(nextFatCount - 109, 127);
                    if (nextFatCount == fatSectorCount && nextDifatCount == difatSectorCount) {
                        break;
                    }

                    fatSectorCount = nextFatCount;
                    difatSectorCount = nextDifatCount;
                }

                var sectors = new List<byte[]>();
                foreach (var chain in regularChains) {
                    chain.StartSector = sectors.Count;
                    sectors.AddRange(chain.Sectors);
                    chain.Info.StartSector = (uint)chain.StartSector;
                }

                int miniStreamStart = sectors.Count;
                sectors.AddRange(miniStreamSectors);

                int miniFatStart = sectors.Count;
                sectors.AddRange(miniFatSectors);

                int directoryStart = sectors.Count;
                for (int i = 0; i < directorySectorCount; i++) {
                    sectors.Add(new byte[SectorSize]);
                }

                var directorySectors = SplitIntoSectors(BuildDirectory(streamInfos, miniStreamBytes.Length, miniStreamStart));
                for (int i = 0; i < directorySectors.Count; i++) {
                    sectors[directoryStart + i] = directorySectors[i];
                }

                int fatStart = sectors.Count;
                for (int i = 0; i < fatSectorCount; i++) {
                    sectors.Add(new byte[SectorSize]);
                }

                int difatStart = sectors.Count;
                for (int i = 0; i < difatSectorCount; i++) {
                    sectors.Add(new byte[SectorSize]);
                }

                uint[] fat = Enumerable.Repeat(FreeSect, sectors.Count).ToArray();
                foreach (var chain in regularChains) {
                    MarkChain(fat, chain.StartSector, chain.Sectors.Count);
                }
                MarkChain(fat, miniStreamStart, miniStreamSectors.Count);
                MarkChain(fat, miniFatStart, miniFatSectors.Count);
                MarkChain(fat, directoryStart, directorySectors.Count);
                for (int i = 0; i < fatSectorCount; i++) fat[fatStart + i] = FatSect;
                for (int i = 0; i < difatSectorCount; i++) fat[difatStart + i] = DifSect;

                WriteFatSectors(sectors, fat, fatStart, fatSectorCount);
                WriteDifatSectors(sectors, fatStart, fatSectorCount, difatStart, difatSectorCount);

                using var output = new MemoryStream(512 + sectors.Count * SectorSize);
                WriteHeader(output, fatStart, fatSectorCount, directoryStart, miniFatStart, miniFatSectors.Count, difatStart, difatSectorCount);
                foreach (var sector in sectors) {
                    output.Write(sector, 0, sector.Length);
                }

                return output.ToArray();
            }

            public static bool TryRead(byte[] bytes, out Dictionary<string, byte[]> streams) {
                streams = new Dictionary<string, byte[]>(StringComparer.Ordinal);
                try {
                    if (bytes.Length < SectorSize) return false;
                    byte[] signature = { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 };
                    for (int i = 0; i < signature.Length; i++) {
                        if (bytes[i] != signature[i]) return false;
                    }

                    int sectorShift = ReadUInt16(bytes, 30);
                    int sectorSize = 1 << sectorShift;
                    if (sectorSize != SectorSize) return false;

                    int fatSectorCount = (int)ReadUInt32(bytes, 44);
                    uint directoryStart = ReadUInt32(bytes, 48);
                    uint miniCutoff = ReadUInt32(bytes, 56);
                    uint miniFatStart = ReadUInt32(bytes, 60);
                    int miniFatSectorCount = (int)ReadUInt32(bytes, 64);
                    uint firstDifat = ReadUInt32(bytes, 68);
                    int difatSectorCount = (int)ReadUInt32(bytes, 72);

                    List<uint> fatSectorIds = ReadDifat(bytes, firstDifat, difatSectorCount, fatSectorCount);
                    uint[] fat = ReadFat(bytes, fatSectorIds);
                    byte[] directoryBytes = ReadRegularStream(bytes, fat, directoryStart, long.MaxValue);
                    var entries = ReadDirectoryEntries(directoryBytes);
                    var root = entries.FirstOrDefault(e => e.ObjectType == 5);
                    byte[] miniStream = root == null || root.StartSector == EndOfChain
                        ? Array.Empty<byte>()
                        : ReadRegularStream(bytes, fat, root.StartSector, root.Size);
                    uint[] miniFat = miniFatStart == EndOfChain || miniFatSectorCount == 0
                        ? Array.Empty<uint>()
                        : BytesToUInt32Array(ReadRegularStream(bytes, fat, miniFatStart, miniFatSectorCount * SectorSize));

                    foreach (var entry in entries.Where(e => e.ObjectType == 2)) {
                        byte[] data = entry.Size < miniCutoff
                            ? ReadMiniStream(miniStream, miniFat, entry.StartSector, entry.Size)
                            : ReadRegularStream(bytes, fat, entry.StartSector, entry.Size);
                        streams[entry.Name] = data;
                    }

                    return true;
                } catch {
                    streams = new Dictionary<string, byte[]>(StringComparer.Ordinal);
                    return false;
                }
            }

            private static byte[] BuildMiniStream(List<StreamInfo> miniStreams) {
                using var output = new MemoryStream();
                uint nextMiniSector = 0;
                foreach (var stream in miniStreams) {
                    stream.StartSector = nextMiniSector;
                    int sectors = CeilingDiv(stream.Data.Length, MiniSectorSize);
                    byte[] padded = new byte[sectors * MiniSectorSize];
                    Buffer.BlockCopy(stream.Data, 0, padded, 0, stream.Data.Length);
                    output.Write(padded, 0, padded.Length);
                    nextMiniSector += (uint)sectors;
                }

                return output.ToArray();
            }

            private static byte[] BuildMiniFat(List<StreamInfo> miniStreams) {
                var entries = new List<uint>();
                uint current = 0;
                foreach (var stream in miniStreams) {
                    int sectors = CeilingDiv(stream.Data.Length, MiniSectorSize);
                    for (int i = 0; i < sectors; i++) {
                        entries.Add(i == sectors - 1 ? EndOfChain : current + 1);
                        current++;
                    }
                }

                using var output = new MemoryStream();
                foreach (uint entry in entries) {
                    OfficeEncryption.WriteUInt32(output, entry);
                }

                return output.ToArray();
            }

            private static byte[] BuildDirectory(List<StreamInfo> streams, int miniStreamSize, int miniStreamStart) {
                using var output = new MemoryStream();
                WriteDirectoryEntry(output, "Root Entry", 5, EndOfChain, EndOfChain, streams.Count > 0 ? 1u : EndOfChain, miniStreamSize > 0 ? (uint)miniStreamStart : EndOfChain, (ulong)miniStreamSize);
                for (int i = 0; i < streams.Count; i++) {
                    uint left = EndOfChain;
                    uint right = i + 1 < streams.Count ? (uint)(i + 2) : EndOfChain;
                    WriteDirectoryEntry(output, streams[i].Name, 2, left, right, EndOfChain, streams[i].StartSector, (ulong)streams[i].Data.Length);
                }

                int remainder = (int)(output.Length % SectorSize);
                if (remainder != 0) {
                    output.Write(new byte[SectorSize - remainder], 0, SectorSize - remainder);
                }

                return output.ToArray();
            }

            private static List<byte[]> SplitIntoSectors(byte[] data) {
                if (data.Length == 0) return new List<byte[]>();
                int count = CeilingDiv(data.Length, SectorSize);
                var sectors = new List<byte[]>(count);
                for (int i = 0; i < count; i++) {
                    byte[] sector = new byte[SectorSize];
                    int offset = i * SectorSize;
                    int length = Math.Min(SectorSize, data.Length - offset);
                    Buffer.BlockCopy(data, offset, sector, 0, length);
                    sectors.Add(sector);
                }

                return sectors;
            }

            private static void MarkChain(uint[] fat, int start, int count) {
                if (count <= 0 || start < 0) return;
                for (int i = 0; i < count; i++) {
                    fat[start + i] = i == count - 1 ? EndOfChain : (uint)(start + i + 1);
                }
            }

            private static void WriteFatSectors(List<byte[]> sectors, uint[] fat, int fatStart, int fatSectorCount) {
                int index = 0;
                for (int i = 0; i < fatSectorCount; i++) {
                    using var stream = new MemoryStream(sectors[fatStart + i], writable: true);
                    for (int j = 0; j < SectorSize / 4; j++) {
                        OfficeEncryption.WriteUInt32(stream, index < fat.Length ? fat[index++] : FreeSect);
                    }
                }
            }

            private static void WriteDifatSectors(List<byte[]> sectors, int fatStart, int fatSectorCount, int difatStart, int difatSectorCount) {
                int fatIndex = 109;
                for (int i = 0; i < difatSectorCount; i++) {
                    using var stream = new MemoryStream(sectors[difatStart + i], writable: true);
                    for (int j = 0; j < 127; j++) {
                        if (fatIndex < fatSectorCount) {
                            OfficeEncryption.WriteUInt32(stream, (uint)(fatStart + fatIndex));
                            fatIndex++;
                        } else {
                            OfficeEncryption.WriteUInt32(stream, FreeSect);
                        }
                    }
                    OfficeEncryption.WriteUInt32(stream, i + 1 < difatSectorCount ? (uint)(difatStart + i + 1) : EndOfChain);
                }
            }

            private static void WriteHeader(Stream output, int fatStart, int fatSectorCount, int directoryStart, int miniFatStart, int miniFatSectorCount, int difatStart, int difatSectorCount) {
                byte[] header = new byte[SectorSize];
                using var stream = new MemoryStream(header, writable: true);
                stream.Write(new byte[] { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 }, 0, 8);
                stream.Position = 24;
                OfficeEncryption.WriteUInt16(stream, 0x003e);
                OfficeEncryption.WriteUInt16(stream, 0x0003);
                OfficeEncryption.WriteUInt16(stream, 0xfffe);
                OfficeEncryption.WriteUInt16(stream, 0x0009);
                OfficeEncryption.WriteUInt16(stream, 0x0006);
                stream.Position = 44;
                OfficeEncryption.WriteUInt32(stream, (uint)fatSectorCount);
                OfficeEncryption.WriteUInt32(stream, (uint)directoryStart);
                OfficeEncryption.WriteUInt32(stream, 0);
                OfficeEncryption.WriteUInt32(stream, MiniStreamCutoff);
                OfficeEncryption.WriteUInt32(stream, miniFatSectorCount > 0 ? (uint)miniFatStart : EndOfChain);
                OfficeEncryption.WriteUInt32(stream, (uint)miniFatSectorCount);
                OfficeEncryption.WriteUInt32(stream, difatSectorCount > 0 ? (uint)difatStart : EndOfChain);
                OfficeEncryption.WriteUInt32(stream, (uint)difatSectorCount);
                for (int i = 0; i < 109; i++) {
                    OfficeEncryption.WriteUInt32(stream, i < fatSectorCount && i < 109 ? (uint)(fatStart + i) : FreeSect);
                }
                output.Write(header, 0, header.Length);
            }

            private static void WriteDirectoryEntry(Stream output, string name, byte objectType, uint leftSibling, uint rightSibling, uint childId, uint startSector, ulong size) {
                byte[] entry = new byte[DirectoryEntrySize];
                byte[] nameBytes = Encoding.Unicode.GetBytes(name + '\0');
                if (nameBytes.Length > 64) {
                    throw new InvalidOperationException("Compound file directory entry name is too long.");
                }
                Buffer.BlockCopy(nameBytes, 0, entry, 0, nameBytes.Length);
                WriteUInt16(entry, 64, (ushort)nameBytes.Length);
                entry[66] = objectType;
                entry[67] = 1;
                WriteUInt32(entry, 68, leftSibling);
                WriteUInt32(entry, 72, rightSibling);
                WriteUInt32(entry, 76, childId);
                WriteUInt32(entry, 116, startSector);
                WriteUInt64(entry, 120, size);
                output.Write(entry, 0, entry.Length);
            }

            private static List<uint> ReadDifat(byte[] bytes, uint firstDifat, int difatSectorCount, int fatSectorCount) {
                var result = new List<uint>(fatSectorCount);
                for (int i = 0; i < 109 && result.Count < fatSectorCount; i++) {
                    uint sector = ReadUInt32(bytes, 76 + i * 4);
                    if (sector != FreeSect) result.Add(sector);
                }

                uint next = firstDifat;
                for (int d = 0; d < difatSectorCount && next != EndOfChain && result.Count < fatSectorCount; d++) {
                    int offset = SectorOffset(next);
                    for (int i = 0; i < 127 && result.Count < fatSectorCount; i++) {
                        uint sector = ReadUInt32(bytes, offset + i * 4);
                        if (sector != FreeSect) result.Add(sector);
                    }
                    next = ReadUInt32(bytes, offset + 127 * 4);
                }

                return result;
            }

            private static uint[] ReadFat(byte[] bytes, List<uint> fatSectorIds) {
                var entries = new List<uint>(fatSectorIds.Count * (SectorSize / 4));
                foreach (uint sector in fatSectorIds) {
                    int offset = SectorOffset(sector);
                    for (int i = 0; i < SectorSize / 4; i++) {
                        entries.Add(ReadUInt32(bytes, offset + i * 4));
                    }
                }
                return entries.ToArray();
            }

            private static byte[] ReadRegularStream(byte[] bytes, uint[] fat, uint startSector, long size) {
                if (startSector == EndOfChain) return Array.Empty<byte>();

                using var output = new MemoryStream();
                uint sector = startSector;
                var visited = new HashSet<uint>();
                while (sector != EndOfChain && sector != FreeSect) {
                    if (sector >= fat.Length || !visited.Add(sector)) {
                        throw new InvalidDataException("Invalid compound file sector chain.");
                    }

                    int offset = SectorOffset(sector);
                    output.Write(bytes, offset, SectorSize);
                    sector = fat[sector];
                    if (output.Length >= size && size != long.MaxValue) {
                        break;
                    }
                }

                byte[] data = output.ToArray();
                if (size != long.MaxValue && data.LongLength > size) {
                    Array.Resize(ref data, (int)size);
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
                    output.Write(miniStream, offset, Math.Min(MiniSectorSize, miniStream.Length - offset));
                    sector = miniFat[sector];
                    if (output.Length >= size) {
                        break;
                    }
                }

                byte[] data = output.ToArray();
                if (data.LongLength > size) {
                    Array.Resize(ref data, (int)size);
                }
                return data;
            }

            private static List<DirectoryEntry> ReadDirectoryEntries(byte[] directoryBytes) {
                var result = new List<DirectoryEntry>();
                for (int offset = 0; offset + DirectoryEntrySize <= directoryBytes.Length; offset += DirectoryEntrySize) {
                    ushort nameLength = ReadUInt16(directoryBytes, offset + 64);
                    byte objectType = directoryBytes[offset + 66];
                    if (objectType == 0 || nameLength < 2 || nameLength > 64) continue;
                    string name = Encoding.Unicode.GetString(directoryBytes, offset, nameLength - 2);
                    result.Add(new DirectoryEntry(
                        name,
                        objectType,
                        ReadUInt32(directoryBytes, offset + 116),
                        (long)ReadUInt64(directoryBytes, offset + 120)));
                }
                return result;
            }

            private static uint[] BytesToUInt32Array(byte[] bytes) {
                uint[] result = new uint[bytes.Length / 4];
                for (int i = 0; i < result.Length; i++) {
                    result[i] = ReadUInt32(bytes, i * 4);
                }
                return result;
            }

            private static int SectorOffset(uint sector) {
                return checked(SectorSize + (int)sector * SectorSize);
            }

            private static int CeilingDiv(int value, int divisor) {
                return value == 0 ? 0 : ((value - 1) / divisor) + 1;
            }

            private static void WriteUInt16(byte[] buffer, int offset, ushort value) {
                buffer[offset] = (byte)(value & 0xff);
                buffer[offset + 1] = (byte)((value >> 8) & 0xff);
            }

            private static void WriteUInt32(byte[] buffer, int offset, uint value) {
                buffer[offset] = (byte)(value & 0xff);
                buffer[offset + 1] = (byte)((value >> 8) & 0xff);
                buffer[offset + 2] = (byte)((value >> 16) & 0xff);
                buffer[offset + 3] = (byte)((value >> 24) & 0xff);
            }

            private static void WriteUInt64(byte[] buffer, int offset, ulong value) {
                WriteUInt32(buffer, offset, (uint)(value & 0xffffffff));
                WriteUInt32(buffer, offset + 4, (uint)(value >> 32));
            }

            private sealed class StreamInfo {
                public StreamInfo(string name, byte[] data) {
                    Name = name;
                    Data = data;
                }

                public string Name { get; }
                public byte[] Data { get; }
                public uint StartSector { get; set; } = EndOfChain;
            }

            private sealed class RegularChain {
                public RegularChain(StreamInfo info, List<byte[]> sectors) {
                    Info = info;
                    Sectors = sectors;
                }

                public StreamInfo Info { get; }
                public List<byte[]> Sectors { get; }
                public int StartSector { get; set; } = -1;
            }

            private sealed class DirectoryEntry {
                public DirectoryEntry(string name, byte objectType, uint startSector, long size) {
                    Name = name;
                    ObjectType = objectType;
                    StartSector = startSector;
                    Size = size;
                }

                public string Name { get; }
                public byte ObjectType { get; }
                public uint StartSector { get; }
                public long Size { get; }
            }
        }
    }
}
