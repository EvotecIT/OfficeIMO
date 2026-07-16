using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeIMO.Drawing.Internal {
    /// <summary>
    /// Writes simple OLE compound document containers used by legacy Office binary formats.
    /// </summary>
    internal static class OfficeCompoundFileWriter {
        private const int SectorSize = 512;
        private const int MiniSectorSize = 64;
        private const int MiniStreamCutoffSize = 4096;
        private const uint FreeSect = 0xffffffff;
        private const uint EndOfChain = 0xfffffffe;
        private const uint FatSect = 0xfffffffd;
        private const uint DifSect = 0xfffffffc;

        internal static byte[] Write(IReadOnlyList<OfficeCompoundStream> streams, Guid? rootClassId = null) {
            if (streams == null) throw new ArgumentNullException(nameof(streams));
            if (streams.Count == 0) throw new ArgumentException("At least one compound stream is required.", nameof(streams));

            return Write(OfficeCompoundWriterLayout.Create(streams), rootClassId);
        }

        /// <summary>Rewrites selected streams while retaining the source directory hierarchy and metadata.</summary>
        internal static byte[] Rewrite(OfficeCompoundFile source,
            IReadOnlyDictionary<string, byte[]> replacementStreams,
            IReadOnlyCollection<string>? removedPaths = null) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (replacementStreams == null) throw new ArgumentNullException(nameof(replacementStreams));
            var removals = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (removedPaths != null) {
                foreach (string path in removedPaths) {
                    if (path == null) throw new ArgumentException("Removed compound paths cannot be null.", nameof(removedPaths));
                    string normalized = NormalizePath(path);
                    if (normalized.Length == 0) {
                        throw new ArgumentException("The compound root cannot be removed.", nameof(removedPaths));
                    }
                    removals.Add(normalized);
                }
            }
            var replacements = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, byte[]> replacement in replacementStreams) {
                string normalized = NormalizePath(replacement.Key);
                if (IsRemoved(normalized, removals)) {
                    throw new ArgumentException(
                        $"Replacement stream '{normalized}' is inside a removed compound path.",
                        nameof(replacementStreams));
                }
                replacements[normalized] = replacement.Value
                    ?? throw new ArgumentException("Replacement stream bytes cannot be null.", nameof(replacementStreams));
            }

            var streams = new List<OfficeCompoundStream>();
            var retainedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (OfficeCompoundFileEntry entry in source.Entries.Where(entry => entry.IsStream && !entry.IsFallback)) {
                if (IsRemoved(entry.Path, removals)) continue;
                byte[] bytes = replacements.TryGetValue(entry.Path, out byte[]? replacement)
                    ? replacement
                    : source.Streams[entry.Path];
                streams.Add(new OfficeCompoundStream(entry.Path, bytes));
                retainedPaths.Add(entry.Path);
            }
            foreach (KeyValuePair<string, byte[]> replacement in replacements) {
                if (retainedPaths.Add(replacement.Key)) {
                    streams.Add(new OfficeCompoundStream(replacement.Key, replacement.Value));
                }
            }
            if (streams.Count == 0) {
                throw new ArgumentException("At least one compound stream is required.", nameof(source));
            }
            return Write(OfficeCompoundWriterLayout.Create(streams, source, removals),
                source.RootEntry.ClassId == Guid.Empty ? null : source.RootEntry.ClassId);
        }

        private static string NormalizePath(string path) =>
            path.Replace('\\', '/').Trim('/');

        private static bool IsRemoved(string path, IReadOnlyCollection<string> removals) {
            foreach (string removal in removals) {
                if (string.Equals(path, removal, StringComparison.OrdinalIgnoreCase)
                    || path.StartsWith(removal + "/", StringComparison.OrdinalIgnoreCase)) {
                    return true;
                }
            }
            return false;
        }

        private static byte[] Write(OfficeCompoundWriterLayout directoryLayout, Guid? rootClassId) {

            PaddedStream[] paddedStreams = directoryLayout.Streams
                .Select(PadStream)
                .OrderBy(stream => stream.Entry.Path, StringComparer.OrdinalIgnoreCase)
                .ToArray();

            MiniStreamLayout miniStreamLayout = MiniStreamLayout.Create(paddedStreams);
            int regularStreamSectorCount = 0;
            foreach (PaddedStream stream in paddedStreams) {
                if (stream.IsMiniStream) {
                    continue;
                }

                stream.StartSector = unchecked((uint)regularStreamSectorCount);
                regularStreamSectorCount += stream.PaddedBytes.Length / SectorSize;
            }

            uint miniStreamStartSector = EndOfChain;
            if (miniStreamLayout.StreamBytes.Length > 0) {
                miniStreamStartSector = unchecked((uint)regularStreamSectorCount);
                regularStreamSectorCount += miniStreamLayout.StreamBytes.Length / SectorSize;
            }

            int directorySectorCount = CalculateDirectorySectorCount(directoryLayout.Entries.Count);
            int miniFatSectorCount = miniStreamLayout.FatBytes.Length / SectorSize;
            int sectorCountBeforeFat = regularStreamSectorCount + directorySectorCount + miniFatSectorCount;
            CalculateAllocationTableSectorCounts(sectorCountBeforeFat, out int fatSectorCount, out int difatSectorCount);

            int directorySector = regularStreamSectorCount;
            uint miniFatStartSector = miniFatSectorCount == 0
                ? EndOfChain
                : unchecked((uint)(directorySector + directorySectorCount));
            int firstFatSector = directorySector + directorySectorCount;
            if (miniFatSectorCount > 0) {
                firstFatSector += miniFatSectorCount;
            }
            int firstDifatSector = firstFatSector + fatSectorCount;

            byte[] directory = BuildDirectory(directoryLayout, paddedStreams, directorySectorCount,
                miniStreamStartSector, miniStreamLayout.StreamLength, rootClassId);
            byte[] fat = BuildFat(
                paddedStreams,
                miniStreamStartSector,
                miniStreamLayout.StreamBytes.Length / SectorSize,
                directorySector,
                directorySectorCount,
                miniFatStartSector,
                miniFatSectorCount,
                firstFatSector,
                fatSectorCount,
                firstDifatSector,
                difatSectorCount);
            byte[] difat = BuildDifat(firstFatSector, fatSectorCount, firstDifatSector, difatSectorCount);

            using var output = new MemoryStream();
            output.Write(BuildHeader(directorySector, firstFatSector, fatSectorCount, miniFatStartSector,
                miniFatSectorCount, firstDifatSector, difatSectorCount), 0, SectorSize);
            foreach (PaddedStream stream in paddedStreams) {
                if (!stream.IsMiniStream) {
                    output.Write(stream.PaddedBytes, 0, stream.PaddedBytes.Length);
                }
            }

            if (miniStreamLayout.StreamBytes.Length > 0) {
                output.Write(miniStreamLayout.StreamBytes, 0, miniStreamLayout.StreamBytes.Length);
            }

            output.Write(directory, 0, directory.Length);
            if (miniStreamLayout.FatBytes.Length > 0) {
                output.Write(miniStreamLayout.FatBytes, 0, miniStreamLayout.FatBytes.Length);
            }

            output.Write(fat, 0, fat.Length);
            if (difat.Length > 0) output.Write(difat, 0, difat.Length);
            return output.ToArray();
        }

        internal static long GetSerializedLength(IReadOnlyList<OfficeCompoundStream> streams) {
            if (streams == null) throw new ArgumentNullException(nameof(streams));
            if (streams.Count == 0) throw new ArgumentException("At least one compound stream is required.", nameof(streams));

            OfficeCompoundWriterLayout layout = OfficeCompoundWriterLayout.Create(streams);
            int regularStreamSectorCount = 0;
            int miniSectorCount = 0;
            foreach (OfficeCompoundWriterEntry stream in layout.Streams) {
                int length = stream.Bytes?.Length ?? 0;
                if (length == 0) continue;
                if (length < MiniStreamCutoffSize) {
                    miniSectorCount = checked(miniSectorCount + ((length + MiniSectorSize - 1) / MiniSectorSize));
                } else {
                    regularStreamSectorCount = checked(regularStreamSectorCount +
                        ((length + SectorSize - 1) / SectorSize));
                }
            }

            int miniStreamSectorCount = checked(((miniSectorCount * MiniSectorSize) + SectorSize - 1) / SectorSize);
            int directorySectorCount = CalculateDirectorySectorCount(layout.Entries.Count);
            int miniFatSectorCount = checked(((miniSectorCount * 4) + SectorSize - 1) / SectorSize);
            int sectorCountBeforeFat = checked(regularStreamSectorCount + miniStreamSectorCount +
                directorySectorCount + miniFatSectorCount);
            CalculateAllocationTableSectorCounts(sectorCountBeforeFat, out int fatSectorCount,
                out int difatSectorCount);
            return checked((1L + sectorCountBeforeFat + fatSectorCount + difatSectorCount) * SectorSize);
        }

        private static int CalculateDirectorySectorCount(int directoryEntryCount) {
            return Math.Max(1, (checked(directoryEntryCount * 128) + SectorSize - 1) / SectorSize);
        }

        private static void CalculateAllocationTableSectorCounts(int sectorCountBeforeFat,
            out int fatSectorCount, out int difatSectorCount) {
            fatSectorCount = 1;
            difatSectorCount = 0;
            while (true) {
                int totalSectors = checked(sectorCountBeforeFat + fatSectorCount + difatSectorCount);
                int requiredFatSectors = (totalSectors + 127) / 128;
                int requiredDifatSectors = requiredFatSectors <= 109 ? 0 : (requiredFatSectors - 109 + 126) / 127;
                if (requiredFatSectors == fatSectorCount && requiredDifatSectors == difatSectorCount) return;
                fatSectorCount = requiredFatSectors;
                difatSectorCount = requiredDifatSectors;
            }
        }

        private static byte[] BuildHeader(int directorySector, int firstFatSector, int fatSectorCount,
            uint miniFatStartSector, int miniFatSectorCount, int firstDifatSector, int difatSectorCount) {
            byte[] header = new byte[SectorSize];
            byte[] signature = { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 };
            Buffer.BlockCopy(signature, 0, header, 0, signature.Length);
            WriteUInt16(header, 24, 0x003e);
            WriteUInt16(header, 26, 0x0003);
            WriteUInt16(header, 28, 0xfffe);
            WriteUInt16(header, 30, 0x0009);
            WriteUInt16(header, 32, 0x0006);
            WriteUInt32(header, 44, unchecked((uint)fatSectorCount));
            WriteUInt32(header, 48, unchecked((uint)directorySector));
            WriteUInt32(header, 56, MiniStreamCutoffSize);
            WriteUInt32(header, 60, miniFatStartSector);
            WriteUInt32(header, 64, unchecked((uint)miniFatSectorCount));
            WriteUInt32(header, 68, difatSectorCount == 0 ? EndOfChain : unchecked((uint)firstDifatSector));
            WriteUInt32(header, 72, unchecked((uint)difatSectorCount));

            for (int i = 0; i < 109; i++) {
                uint value = i < fatSectorCount ? unchecked((uint)(firstFatSector + i)) : FreeSect;
                WriteUInt32(header, 76 + i * 4, value);
            }

            return header;
        }

        private static byte[] BuildDirectory(OfficeCompoundWriterLayout layout, IReadOnlyList<PaddedStream> streams,
            int directorySectorCount, uint miniStreamStartSector, int miniStreamLength, Guid? rootClassId) {
            byte[] directory = new byte[checked(directorySectorCount * SectorSize)];
            var paddedByEntry = streams.ToDictionary(stream => stream.Entry);
            foreach (OfficeCompoundWriterEntry entry in layout.Entries) {
                uint startSector = EndOfChain;
                ulong size = 0;
                if (entry.ObjectType == 5) {
                    startSector = miniStreamStartSector;
                    size = unchecked((ulong)miniStreamLength);
                } else if (entry.ObjectType == 2) {
                    PaddedStream stream = paddedByEntry[entry];
                    startSector = stream.StartSector;
                    size = unchecked((ulong)stream.OriginalLength);
                }
                WriteDirectoryEntry(
                    directory,
                    checked(entry.DirectoryIndex * 128),
                    entry.Name,
                    entry.ObjectType,
                    entry.LeftSiblingId,
                    entry.RightSiblingId,
                    entry.ChildId,
                    startSector,
                    size,
                    entry.ObjectType == 5 ? rootClassId ?? entry.ClassId : entry.ClassId,
                    entry.StateBits,
                    entry.CreationTime,
                    entry.ModifiedTime);
            }

            return directory;
        }

        private static byte[] BuildFat(
            IReadOnlyList<PaddedStream> streams,
            uint miniStreamStartSector,
            int miniStreamSectorCount,
            int directorySector,
            int directorySectorCount,
            uint miniFatStartSector,
            int miniFatSectorCount,
            int firstFatSector,
            int fatSectorCount,
            int firstDifatSector,
            int difatSectorCount) {
            byte[] fat = new byte[checked(fatSectorCount * SectorSize)];
            for (int offset = 0; offset < fat.Length; offset += 4) {
                WriteUInt32(fat, offset, FreeSect);
            }

            foreach (PaddedStream stream in streams) {
                if (!stream.IsMiniStream) {
                    WriteFatChain(fat, stream.StartSector, stream.PaddedBytes.Length / SectorSize);
                }
            }

            if (miniStreamSectorCount > 0) {
                WriteFatChain(fat, miniStreamStartSector, miniStreamSectorCount);
            }

            WriteFatChain(fat, unchecked((uint)directorySector), directorySectorCount);
            if (miniFatSectorCount > 0) {
                WriteFatChain(fat, miniFatStartSector, miniFatSectorCount);
            }

            for (int i = 0; i < fatSectorCount; i++) {
                WriteFatEntry(fat, firstFatSector + i, FatSect);
            }
            for (int i = 0; i < difatSectorCount; i++) {
                WriteFatEntry(fat, firstDifatSector + i, DifSect);
            }

            return fat;
        }

        private static byte[] BuildDifat(int firstFatSector, int fatSectorCount, int firstDifatSector,
            int difatSectorCount) {
            if (difatSectorCount == 0) return Array.Empty<byte>();
            byte[] difat = new byte[checked(difatSectorCount * SectorSize)];
            for (int offset = 0; offset < difat.Length; offset += 4) WriteUInt32(difat, offset, FreeSect);
            int fatIndex = 109;
            for (int sectorIndex = 0; sectorIndex < difatSectorCount; sectorIndex++) {
                int offset = sectorIndex * SectorSize;
                for (int entryIndex = 0; entryIndex < 127 && fatIndex < fatSectorCount; entryIndex++, fatIndex++) {
                    WriteUInt32(difat, offset + entryIndex * 4, unchecked((uint)(firstFatSector + fatIndex)));
                }
                uint next = sectorIndex + 1 == difatSectorCount
                    ? EndOfChain
                    : unchecked((uint)(firstDifatSector + sectorIndex + 1));
                WriteUInt32(difat, offset + 127 * 4, next);
            }
            return difat;
        }

        private static void WriteFatChain(byte[] fat, uint firstSector, int sectorCount) {
            if (sectorCount == 0 || firstSector == EndOfChain) {
                return;
            }

            for (int i = 0; i < sectorCount; i++) {
                bool lastSector = i + 1 == sectorCount;
                uint sector = unchecked(firstSector + (uint)i);
                WriteFatEntry(fat, checked((int)sector), lastSector ? EndOfChain : unchecked(sector + 1));
            }
        }

        private static void WriteDirectoryEntry(byte[] buffer, int offset, string name, byte type, uint left, uint right,
            uint child, uint startSector, ulong size, Guid? classId, uint stateBits, ulong creationTime,
            ulong modifiedTime) {
            byte[] nameBytes = Encoding.Unicode.GetBytes(name + '\0');
            Buffer.BlockCopy(nameBytes, 0, buffer, offset, nameBytes.Length);
            WriteUInt16(buffer, offset + 64, checked((ushort)nameBytes.Length));
            buffer[offset + 66] = type;
            buffer[offset + 67] = 1;
            WriteUInt32(buffer, offset + 68, left);
            WriteUInt32(buffer, offset + 72, right);
            WriteUInt32(buffer, offset + 76, child);
            if (classId.HasValue) {
                byte[] classIdBytes = classId.Value.ToByteArray();
                Buffer.BlockCopy(classIdBytes, 0, buffer, offset + 80, classIdBytes.Length);
            }
            WriteUInt32(buffer, offset + 96, stateBits);
            WriteUInt64(buffer, offset + 100, creationTime);
            WriteUInt64(buffer, offset + 108, modifiedTime);
            WriteUInt32(buffer, offset + 116, startSector);
            WriteUInt64(buffer, offset + 120, size);
        }

        private static byte[] PadToSector(byte[] data) {
            int paddedLength = ((data.Length + SectorSize - 1) / SectorSize) * SectorSize;
            if (paddedLength == data.Length) {
                return data;
            }

            byte[] padded = new byte[paddedLength];
            Buffer.BlockCopy(data, 0, padded, 0, data.Length);
            return padded;
        }

        private static byte[] PadToMiniSector(byte[] data) {
            int paddedLength = ((data.Length + MiniSectorSize - 1) / MiniSectorSize) * MiniSectorSize;
            if (paddedLength == data.Length) {
                return data;
            }

            byte[] padded = new byte[paddedLength];
            Buffer.BlockCopy(data, 0, padded, 0, data.Length);
            return padded;
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
            WriteUInt32(buffer, offset, unchecked((uint)(value & 0xffffffffUL)));
            WriteUInt32(buffer, offset + 4, unchecked((uint)(value >> 32)));
        }

        private static void WriteFatEntry(byte[] fat, int sector, uint value) {
            WriteUInt32(fat, checked(sector * 4), value);
        }

        private static PaddedStream PadStream(OfficeCompoundWriterEntry entry) {
            byte[] bytes = entry.Bytes ?? Array.Empty<byte>();
            bool isMiniStream = bytes.Length < MiniStreamCutoffSize;
            byte[] paddedBytes = isMiniStream ? PadToMiniSector(bytes) : PadToSector(bytes);
            return new PaddedStream(entry, bytes.Length, paddedBytes, isMiniStream);
        }

        private sealed class PaddedStream {
            internal PaddedStream(OfficeCompoundWriterEntry entry, int originalLength, byte[] paddedBytes, bool isMiniStream) {
                Entry = entry;
                OriginalLength = originalLength;
                PaddedBytes = paddedBytes;
                IsMiniStream = isMiniStream;
                StartSector = EndOfChain;
            }

            internal OfficeCompoundWriterEntry Entry { get; }

            internal int OriginalLength { get; }

            internal byte[] PaddedBytes { get; }

            internal bool IsMiniStream { get; }

            internal uint StartSector { get; set; }
        }

        private sealed class MiniStreamLayout {
            private MiniStreamLayout(byte[] streamBytes, byte[] fatBytes, int streamLength) {
                StreamBytes = streamBytes;
                FatBytes = fatBytes;
                StreamLength = streamLength;
            }

            internal byte[] StreamBytes { get; }

            internal byte[] FatBytes { get; }

            internal int StreamLength { get; }

            internal static MiniStreamLayout Create(IReadOnlyList<PaddedStream> streams) {
                var miniStreams = streams
                    .Where(stream => stream.IsMiniStream && stream.PaddedBytes.Length > 0)
                    .ToArray();
                if (miniStreams.Length == 0) {
                    return new MiniStreamLayout(Array.Empty<byte>(), Array.Empty<byte>(), 0);
                }

                int miniSectorCount = 0;
                using var miniStream = new MemoryStream();
                foreach (PaddedStream stream in miniStreams) {
                    stream.StartSector = unchecked((uint)miniSectorCount);
                    miniSectorCount += stream.PaddedBytes.Length / MiniSectorSize;
                    miniStream.Write(stream.PaddedBytes, 0, stream.PaddedBytes.Length);
                }

                byte[] miniFat = new byte[(((miniSectorCount * 4) + SectorSize - 1) / SectorSize) * SectorSize];
                for (int offset = 0; offset < miniFat.Length; offset += 4) {
                    WriteUInt32(miniFat, offset, FreeSect);
                }

                foreach (PaddedStream stream in miniStreams) {
                    int streamMiniSectorCount = stream.PaddedBytes.Length / MiniSectorSize;
                    for (int i = 0; i < streamMiniSectorCount; i++) {
                        bool lastSector = i + 1 == streamMiniSectorCount;
                        uint sector = unchecked(stream.StartSector + (uint)i);
                        WriteUInt32(miniFat, checked((int)sector * 4), lastSector ? EndOfChain : unchecked(sector + 1));
                    }
                }

                byte[] streamBytes = PadToSector(miniStream.ToArray());
                return new MiniStreamLayout(streamBytes, miniFat, checked(miniSectorCount * MiniSectorSize));
            }
        }

    }

    /// <summary>
    /// Named stream payload to write into an OLE compound document.
    /// </summary>
    internal readonly struct OfficeCompoundStream {
        internal OfficeCompoundStream(string name, byte[] bytes) {
            Name = name ?? throw new ArgumentNullException(nameof(name));
            Bytes = bytes ?? throw new ArgumentNullException(nameof(bytes));
        }

        internal string Name { get; }

        internal byte[] Bytes { get; }
    }
}
