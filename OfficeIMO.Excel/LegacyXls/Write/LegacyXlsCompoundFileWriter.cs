using System.Text;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static class LegacyXlsCompoundFileWriter {
        private const int SectorSize = 512;
        private const int MiniSectorSize = 64;
        private const int MiniStreamCutoffSize = 4096;
        private const uint FreeSect = 0xffffffff;
        private const uint EndOfChain = 0xfffffffe;
        private const uint FatSect = 0xfffffffd;

        internal static byte[] Write(byte[] workbookStream) {
            return Write(workbookStream, Array.Empty<LegacyXlsCompoundStream>());
        }

        internal static byte[] Write(byte[] workbookStream, IReadOnlyList<LegacyXlsCompoundStream> additionalStreams) {
            if (workbookStream == null) throw new ArgumentNullException(nameof(workbookStream));
            if (additionalStreams == null) throw new ArgumentNullException(nameof(additionalStreams));

            var streams = new List<LegacyXlsCompoundStream>(additionalStreams.Count + 1) {
                new LegacyXlsCompoundStream("Workbook", workbookStream)
            };
            streams.AddRange(additionalStreams);

            PaddedStream[] paddedStreams = streams
                .Select(PadStream)
                .OrderBy(stream => stream.Name, DirectoryNameComparer.Instance)
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

            int directorySectorCount = CalculateDirectorySectorCount(paddedStreams.Length + 1);
            int miniFatSectorCount = miniStreamLayout.FatBytes.Length / SectorSize;
            int sectorCountBeforeFat = regularStreamSectorCount + directorySectorCount + miniFatSectorCount;
            int fatSectorCount = CalculateFatSectorCount(sectorCountBeforeFat);
            if (fatSectorCount > 109) {
                throw new NotSupportedException("Native XLS saving currently supports compound files with up to 109 FAT sectors.");
            }

            int directorySector = regularStreamSectorCount;
            uint miniFatStartSector = miniFatSectorCount == 0
                ? EndOfChain
                : unchecked((uint)(directorySector + directorySectorCount));
            int firstFatSector = directorySector + directorySectorCount;
            if (miniFatSectorCount > 0) {
                firstFatSector += miniFatSectorCount;
            }

            byte[] directory = BuildDirectory(paddedStreams, directorySectorCount, miniStreamStartSector, miniStreamLayout.StreamLength);
            byte[] fat = BuildFat(
                paddedStreams,
                miniStreamStartSector,
                miniStreamLayout.StreamBytes.Length / SectorSize,
                directorySector,
                directorySectorCount,
                miniFatStartSector,
                miniFatSectorCount,
                firstFatSector,
                fatSectorCount);

            using var output = new MemoryStream();
            output.Write(BuildHeader(directorySector, firstFatSector, fatSectorCount, miniFatStartSector, miniFatSectorCount), 0, SectorSize);
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
            return output.ToArray();
        }

        private static int CalculateDirectorySectorCount(int directoryEntryCount) {
            return Math.Max(1, (checked(directoryEntryCount * 128) + SectorSize - 1) / SectorSize);
        }

        private static int CalculateFatSectorCount(int sectorCountBeforeFat) {
            int fatSectorCount = 1;
            while (true) {
                int totalSectors = sectorCountBeforeFat + fatSectorCount;
                int requiredFatSectors = (totalSectors + 127) / 128;
                if (requiredFatSectors == fatSectorCount) {
                    return fatSectorCount;
                }

                fatSectorCount = requiredFatSectors;
            }
        }

        private static byte[] BuildHeader(int directorySector, int firstFatSector, int fatSectorCount, uint miniFatStartSector, int miniFatSectorCount) {
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
            WriteUInt32(header, 68, EndOfChain);

            for (int i = 0; i < 109; i++) {
                uint value = i < fatSectorCount ? unchecked((uint)(firstFatSector + i)) : FreeSect;
                WriteUInt32(header, 76 + i * 4, value);
            }

            return header;
        }

        private static byte[] BuildDirectory(IReadOnlyList<PaddedStream> streams, int directorySectorCount, uint miniStreamStartSector, int miniStreamLength) {
            byte[] directory = new byte[checked(directorySectorCount * SectorSize)];
            DirectoryTreeLinks directoryLinks = DirectoryTreeLinks.Create(streams.Count);
            WriteDirectoryEntry(directory, 0, "Root Entry", 5, EndOfChain, EndOfChain, directoryLinks.RootChild, miniStreamStartSector, unchecked((ulong)miniStreamLength));

            for (int i = 0; i < streams.Count; i++) {
                PaddedStream stream = streams[i];
                WriteDirectoryEntry(
                    directory,
                    checked((i + 1) * 128),
                    stream.Name,
                    2,
                    directoryLinks.GetLeftSibling(i),
                    directoryLinks.GetRightSibling(i),
                    EndOfChain,
                    stream.StartSector,
                    unchecked((ulong)stream.OriginalLength));
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
            int fatSectorCount) {
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

            return fat;
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

        private static void WriteDirectoryEntry(byte[] buffer, int offset, string name, byte type, uint left, uint right, uint child, uint startSector, ulong size) {
            byte[] nameBytes = Encoding.Unicode.GetBytes(name + '\0');
            Buffer.BlockCopy(nameBytes, 0, buffer, offset, nameBytes.Length);
            WriteUInt16(buffer, offset + 64, checked((ushort)nameBytes.Length));
            buffer[offset + 66] = type;
            buffer[offset + 67] = 1;
            WriteUInt32(buffer, offset + 68, left);
            WriteUInt32(buffer, offset + 72, right);
            WriteUInt32(buffer, offset + 76, child);
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

        private static PaddedStream PadStream(LegacyXlsCompoundStream stream) {
            if (string.IsNullOrEmpty(stream.Name)) {
                throw new ArgumentException("Compound stream name is required.", nameof(stream));
            }

            bool isMiniStream = stream.Bytes.Length < MiniStreamCutoffSize;
            byte[] paddedBytes = isMiniStream ? PadToMiniSector(stream.Bytes) : PadToSector(stream.Bytes);
            return new PaddedStream(stream.Name, stream.Bytes.Length, paddedBytes, isMiniStream);
        }

        private sealed class PaddedStream {
            internal PaddedStream(string name, int originalLength, byte[] paddedBytes, bool isMiniStream) {
                Name = name;
                OriginalLength = originalLength;
                PaddedBytes = paddedBytes;
                IsMiniStream = isMiniStream;
                StartSector = EndOfChain;
            }

            internal string Name { get; }

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

        private sealed class DirectoryTreeLinks {
            private readonly int[] _left;
            private readonly int[] _right;

            private DirectoryTreeLinks(int[] left, int[] right, int root) {
                _left = left;
                _right = right;
                Root = root;
            }

            internal int Root { get; }

            internal uint RootChild => ToDirectoryEntryId(Root);

            internal static DirectoryTreeLinks Create(int streamCount) {
                var left = Enumerable.Repeat(-1, streamCount).ToArray();
                var right = Enumerable.Repeat(-1, streamCount).ToArray();
                int root = BuildBalancedTree(0, streamCount - 1, left, right);
                return new DirectoryTreeLinks(left, right, root);
            }

            internal uint GetLeftSibling(int streamIndex) {
                return ToDirectoryEntryId(_left[streamIndex]);
            }

            internal uint GetRightSibling(int streamIndex) {
                return ToDirectoryEntryId(_right[streamIndex]);
            }

            private static int BuildBalancedTree(int firstIndex, int lastIndex, int[] left, int[] right) {
                if (firstIndex > lastIndex) {
                    return -1;
                }

                int middleIndex = firstIndex + ((lastIndex - firstIndex) / 2);
                left[middleIndex] = BuildBalancedTree(firstIndex, middleIndex - 1, left, right);
                right[middleIndex] = BuildBalancedTree(middleIndex + 1, lastIndex, left, right);
                return middleIndex;
            }

            private static uint ToDirectoryEntryId(int streamIndex) {
                return streamIndex < 0 ? EndOfChain : unchecked((uint)(streamIndex + 1));
            }
        }

        private sealed class DirectoryNameComparer : IComparer<string> {
            internal static DirectoryNameComparer Instance { get; } = new DirectoryNameComparer();

            public int Compare(string? left, string? right) {
                int ignoreCase = StringComparer.OrdinalIgnoreCase.Compare(left, right);
                return ignoreCase != 0 ? ignoreCase : StringComparer.Ordinal.Compare(left, right);
            }
        }
    }

    internal readonly struct LegacyXlsCompoundStream {
        internal LegacyXlsCompoundStream(string name, byte[] bytes) {
            Name = name ?? throw new ArgumentNullException(nameof(name));
            Bytes = bytes ?? throw new ArgumentNullException(nameof(bytes));
        }

        internal string Name { get; }

        internal byte[] Bytes { get; }
    }
}
