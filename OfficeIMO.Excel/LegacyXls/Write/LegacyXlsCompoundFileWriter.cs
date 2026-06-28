using System.Text;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static class LegacyXlsCompoundFileWriter {
        private const int SectorSize = 512;
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

            PaddedStream[] paddedStreams = streams.Select(PadStream).ToArray();
            int dataSectorCount = paddedStreams.Sum(stream => stream.PaddedBytes.Length / SectorSize);
            int directorySectorCount = CalculateDirectorySectorCount(paddedStreams.Length + 1);
            int fatSectorCount = CalculateFatSectorCount(dataSectorCount, directorySectorCount);
            if (fatSectorCount > 109) {
                throw new NotSupportedException("Native XLS saving currently supports compound files with up to 109 FAT sectors.");
            }

            int directorySector = dataSectorCount;
            int firstFatSector = directorySector + directorySectorCount;
            byte[] directory = BuildDirectory(paddedStreams, directorySectorCount);
            byte[] fat = BuildFat(paddedStreams, dataSectorCount, directorySector, directorySectorCount, firstFatSector, fatSectorCount);

            using var output = new MemoryStream();
            output.Write(BuildHeader(directorySector, firstFatSector, fatSectorCount), 0, SectorSize);
            foreach (PaddedStream stream in paddedStreams) {
                output.Write(stream.PaddedBytes, 0, stream.PaddedBytes.Length);
            }

            output.Write(directory, 0, directory.Length);
            output.Write(fat, 0, fat.Length);
            return output.ToArray();
        }

        private static int CalculateDirectorySectorCount(int directoryEntryCount) {
            return Math.Max(1, (checked(directoryEntryCount * 128) + SectorSize - 1) / SectorSize);
        }

        private static int CalculateFatSectorCount(int dataSectorCount, int directorySectorCount) {
            int fatSectorCount = 1;
            while (true) {
                int totalSectors = dataSectorCount + directorySectorCount + fatSectorCount;
                int requiredFatSectors = (totalSectors + 127) / 128;
                if (requiredFatSectors == fatSectorCount) {
                    return fatSectorCount;
                }

                fatSectorCount = requiredFatSectors;
            }
        }

        private static byte[] BuildHeader(int directorySector, int firstFatSector, int fatSectorCount) {
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
            WriteUInt32(header, 56, 4096);
            WriteUInt32(header, 60, EndOfChain);
            WriteUInt32(header, 68, EndOfChain);

            for (int i = 0; i < 109; i++) {
                uint value = i < fatSectorCount ? unchecked((uint)(firstFatSector + i)) : FreeSect;
                WriteUInt32(header, 76 + i * 4, value);
            }

            return header;
        }

        private static byte[] BuildDirectory(IReadOnlyList<PaddedStream> streams, int directorySectorCount) {
            byte[] directory = new byte[checked(directorySectorCount * SectorSize)];
            WriteDirectoryEntry(directory, 0, "Root Entry", 5, EndOfChain, EndOfChain, streams.Count > 0 ? 1U : EndOfChain, EndOfChain, 0);

            uint startSector = 0;
            for (int i = 0; i < streams.Count; i++) {
                PaddedStream stream = streams[i];
                uint rightSibling = i + 1 < streams.Count ? unchecked((uint)(i + 2)) : EndOfChain;
                WriteDirectoryEntry(
                    directory,
                    checked((i + 1) * 128),
                    stream.Name,
                    2,
                    EndOfChain,
                    rightSibling,
                    EndOfChain,
                    startSector,
                    unchecked((ulong)stream.PaddedBytes.Length));
                startSector += unchecked((uint)(stream.PaddedBytes.Length / SectorSize));
            }

            return directory;
        }

        private static byte[] BuildFat(
            IReadOnlyList<PaddedStream> streams,
            int dataSectorCount,
            int directorySector,
            int directorySectorCount,
            int firstFatSector,
            int fatSectorCount) {
            byte[] fat = new byte[checked(fatSectorCount * SectorSize)];
            for (int offset = 0; offset < fat.Length; offset += 4) {
                WriteUInt32(fat, offset, FreeSect);
            }

            int streamStartSector = 0;
            foreach (PaddedStream stream in streams) {
                int streamSectorCount = stream.PaddedBytes.Length / SectorSize;
                for (int i = 0; i < streamSectorCount; i++) {
                    bool lastSector = i + 1 == streamSectorCount;
                    WriteFatEntry(fat, streamStartSector + i, lastSector ? EndOfChain : unchecked((uint)(streamStartSector + i + 1)));
                }

                streamStartSector += streamSectorCount;
            }

            if (streamStartSector != dataSectorCount) {
                throw new InvalidOperationException("The compound file stream sector count is inconsistent.");
            }

            for (int i = 0; i < directorySectorCount; i++) {
                bool lastDirectorySector = i + 1 == directorySectorCount;
                WriteFatEntry(fat, directorySector + i, lastDirectorySector ? EndOfChain : unchecked((uint)(directorySector + i + 1)));
            }

            for (int i = 0; i < fatSectorCount; i++) {
                WriteFatEntry(fat, firstFatSector + i, FatSect);
            }

            return fat;
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
            int paddedLength = Math.Max(4096, ((data.Length + SectorSize - 1) / SectorSize) * SectorSize);
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

            return new PaddedStream(stream.Name, stream.Bytes.Length, PadToSector(stream.Bytes));
        }

        private sealed class PaddedStream {
            internal PaddedStream(string name, int originalLength, byte[] paddedBytes) {
                Name = name;
                OriginalLength = originalLength;
                PaddedBytes = paddedBytes;
            }

            internal string Name { get; }

            internal int OriginalLength { get; }

            internal byte[] PaddedBytes { get; }
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
