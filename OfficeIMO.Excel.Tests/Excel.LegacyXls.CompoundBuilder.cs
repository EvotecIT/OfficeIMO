using System.Linq;
using System.Text;
using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private static class LegacyXlsCompoundTestBuilder {
            private const int SectorSize = 512;
            private const int MiniSectorSize = 64;
            private const uint FreeSect = 0xffffffff;
            private const uint EndOfChain = 0xfffffffe;
            private const uint FatSect = 0xfffffffd;
            private const uint DifSect = 0xfffffffc;

            internal static byte[] CreateWorkbookCompoundFile(byte[] workbookStream) {
                return CreateWorkbookCompoundFile(workbookStream, includeVbaProjectStorage: false, includeOleObjectStorage: false);
            }

            internal static byte[] CreateWorkbookCompoundFileWithDocumentPropertyStreams(
                byte[] workbookStream,
                byte[] summaryInformationStream,
                byte[] documentSummaryInformationStream) {
                var streams = new[] {
                    new CompoundStreamSpec("Workbook", workbookStream),
                    new CompoundStreamSpec("\u0005SummaryInformation", summaryInformationStream),
                    new CompoundStreamSpec("\u0005DocumentSummaryInformation", documentSummaryInformationStream)
                };
                return CreateRootStreamCompoundFile(streams);
            }

            internal static byte[] CreateWorkbookCompoundFileWithVbaProjectStorage(byte[] workbookStream) {
                return CreateWorkbookCompoundFile(workbookStream, includeVbaProjectStorage: true, includeOleObjectStorage: false);
            }

            internal static byte[] CreateWorkbookCompoundFileWithVbaProjectPayload(
                byte[] workbookStream,
                byte[] vbaPayload) {
                byte[] compoundBytes = CreateWorkbookCompoundFileWithVbaProjectStorage(workbookStream);
                if (!OfficeCompoundFileReader.TryRead(
                    compoundBytes,
                    out OfficeCompoundFile? compoundFile,
                    out string? error)) {
                    throw new InvalidOperationException(error);
                }

                return OfficeCompoundFileWriter.Rewrite(
                    compoundFile!,
                    new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase) {
                        ["_VBA_PROJECT_CUR/VBA/dir"] = vbaPayload
                    });
            }

            internal static byte[] CreateWorkbookCompoundFileWithOleObjectStorage(byte[] workbookStream) {
                return CreateWorkbookCompoundFile(workbookStream, includeVbaProjectStorage: false, includeOleObjectStorage: true);
            }

            internal static byte[] CreateWorkbookCompoundFileWithOleObjectPayload(
                byte[] workbookStream,
                byte[] olePayload) {
                byte[] compoundBytes = CreateWorkbookCompoundFileWithOleObjectStorage(workbookStream);
                if (!OfficeCompoundFileReader.TryRead(
                    compoundBytes,
                    out OfficeCompoundFile? compoundFile,
                    out string? error)) {
                    throw new InvalidOperationException(error);
                }

                return OfficeCompoundFileWriter.Rewrite(
                    compoundFile!,
                    new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase) {
                        ["ObjectPool/OLEPackage/\u0001Ole10Native"] = olePayload
                    });
            }

            internal static byte[] CreateWorkbookCompoundFileWithDigitalSignatureStream(byte[] workbookStream) {
                var streams = new[] {
                    new CompoundStreamSpec("Workbook", workbookStream),
                    new CompoundStreamSpec("_signatures", Encoding.ASCII.GetBytes("synthetic legacy digital signature"))
                };
                return CreateRootStreamCompoundFile(streams);
            }

            internal static byte[] CreateCompoundHeaderWithInvalidSectorChain() {
                byte[] bytes = new byte[SectorSize];
                byte[] signature = { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 };
                Buffer.BlockCopy(signature, 0, bytes, 0, signature.Length);
                bytes[30] = 9;
                bytes[32] = 6;
                return bytes;
            }

            internal static byte[] CreateNonExcelCompoundFile() {
                byte[] streamBytes = PadToRegularStream(Encoding.UTF8.GetBytes("not an Excel workbook stream"));
                byte[] directoryBytes = BuildNonExcelDirectory(streamBytes.Length);
                byte[] fatBytes = BuildFat(streamBytes.Length / SectorSize);

                using var output = new MemoryStream();
                output.Write(BuildHeader(streamBytes.Length / SectorSize), 0, SectorSize);
                output.Write(streamBytes, 0, streamBytes.Length);
                output.Write(directoryBytes, 0, directoryBytes.Length);
                output.Write(fatBytes, 0, fatBytes.Length);
                return output.ToArray();
            }

            internal static byte[] CreateMiniStreamWorkbookCompoundFile(byte[] workbookStream) {
                if (workbookStream.Length >= 4096) {
                    throw new ArgumentException("The workbook stream must be smaller than the compound file mini stream cutoff.", nameof(workbookStream));
                }

                byte[] rootMiniStream = PadToMiniStreamContainer(workbookStream);
                int usedMiniSectorCount = Math.Max(1, (workbookStream.Length + MiniSectorSize - 1) / MiniSectorSize);

                using var output = new MemoryStream();
                output.Write(BuildMiniStreamHeader(), 0, SectorSize);
                output.Write(rootMiniStream, 0, rootMiniStream.Length);
                output.Write(BuildMiniStreamDirectory(workbookStream.Length, rootMiniStream.Length), 0, SectorSize);
                output.Write(BuildMiniFat(usedMiniSectorCount), 0, SectorSize);
                output.Write(BuildMiniStreamFat(), 0, SectorSize);
                return output.ToArray();
            }

            internal static byte[] CreateDifatWorkbookCompoundFile(byte[] workbookStream) {
                const int workbookSectorCount = 8;
                const int dataSectorCount = 13960;
                const int fatSectorCount = 110;

                byte[] workbookBytes = PadToSectorCount(workbookStream, workbookSectorCount);
                int directorySector = dataSectorCount;
                int firstFatSector = directorySector + 1;
                int difatSector = firstFatSector + fatSectorCount;

                using var output = new MemoryStream();
                output.Write(BuildDifatHeader(directorySector, fatSectorCount, firstFatSector, difatSector), 0, SectorSize);
                output.Write(workbookBytes, 0, workbookBytes.Length);

                byte[] emptySector = new byte[SectorSize];
                for (int i = workbookSectorCount; i < dataSectorCount; i++) {
                    output.Write(emptySector, 0, emptySector.Length);
                }

                output.Write(BuildDifatDirectory(workbookBytes.Length), 0, SectorSize);
                byte[] fatBytes = BuildDifatFat(workbookSectorCount, directorySector, firstFatSector, fatSectorCount, difatSector);
                output.Write(fatBytes, 0, fatBytes.Length);
                output.Write(BuildDifatSector(firstFatSector + 109), 0, SectorSize);
                return output.ToArray();
            }

            private static byte[] CreateWorkbookCompoundFile(byte[] workbookStream, bool includeVbaProjectStorage, bool includeOleObjectStorage) {
                byte[] workbookBytes = PadToRegularStream(workbookStream);
                byte[] directoryBytes = BuildDirectory(workbookBytes.Length, includeVbaProjectStorage, includeOleObjectStorage);
                int workbookSectorCount = workbookBytes.Length / SectorSize;
                byte[] fatBytes = BuildFat(workbookSectorCount);

                using var output = new MemoryStream();
                output.Write(BuildHeader(workbookSectorCount), 0, SectorSize);
                output.Write(workbookBytes, 0, workbookBytes.Length);
                output.Write(directoryBytes, 0, directoryBytes.Length);
                output.Write(fatBytes, 0, fatBytes.Length);
                return output.ToArray();
            }

            private static byte[] CreateRootStreamCompoundFile(IReadOnlyList<CompoundStreamSpec> streams) {
                if (streams.Count == 0 || streams.Count > 3) {
                    throw new ArgumentException("The test compound builder supports one to three root streams.", nameof(streams));
                }

                var paddedStreams = streams
                    .Select(stream => {
                        byte[] padded = PadToRegularStream(stream.Bytes);
                        return new CompoundStreamSpec(stream.Name, padded, padded.Length);
                    })
                    .ToArray();
                int dataSectorCount = paddedStreams.Sum(stream => stream.Bytes.Length / SectorSize);
                byte[] directoryBytes = BuildRootStreamDirectory(paddedStreams);
                byte[] fatBytes = BuildRootStreamFat(paddedStreams);

                using var output = new MemoryStream();
                output.Write(BuildHeader(dataSectorCount), 0, SectorSize);
                foreach (CompoundStreamSpec stream in paddedStreams) {
                    output.Write(stream.Bytes, 0, stream.Bytes.Length);
                }

                output.Write(directoryBytes, 0, directoryBytes.Length);
                output.Write(fatBytes, 0, fatBytes.Length);
                return output.ToArray();
            }

            private static byte[] BuildHeader(int workbookSectorCount) {
                byte[] header = new byte[SectorSize];
                byte[] signature = { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 };
                Buffer.BlockCopy(signature, 0, header, 0, signature.Length);
                WriteUInt16(header, 24, 0x003e);
                WriteUInt16(header, 26, 0x0003);
                WriteUInt16(header, 28, 0xfffe);
                WriteUInt16(header, 30, 0x0009);
                WriteUInt16(header, 32, 0x0006);
                WriteUInt32(header, 44, 1);
                WriteUInt32(header, 48, (uint)workbookSectorCount);
                WriteUInt32(header, 56, 4096);
                WriteUInt32(header, 60, EndOfChain);
                WriteUInt32(header, 68, EndOfChain);
                for (int i = 0; i < 109; i++) {
                    WriteUInt32(header, 76 + i * 4, i == 0 ? (uint)(workbookSectorCount + 1) : FreeSect);
                }

                return header;
            }

            private static byte[] BuildDifatHeader(int directorySector, int fatSectorCount, int firstFatSector, int difatSector) {
                byte[] header = new byte[SectorSize];
                byte[] signature = { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 };
                Buffer.BlockCopy(signature, 0, header, 0, signature.Length);
                WriteUInt16(header, 24, 0x003e);
                WriteUInt16(header, 26, 0x0003);
                WriteUInt16(header, 28, 0xfffe);
                WriteUInt16(header, 30, 0x0009);
                WriteUInt16(header, 32, 0x0006);
                WriteUInt32(header, 44, (uint)fatSectorCount);
                WriteUInt32(header, 48, (uint)directorySector);
                WriteUInt32(header, 56, 4096);
                WriteUInt32(header, 60, EndOfChain);
                WriteUInt32(header, 68, (uint)difatSector);
                WriteUInt32(header, 72, 1);
                for (int i = 0; i < 109; i++) {
                    WriteUInt32(header, 76 + i * 4, (uint)(firstFatSector + i));
                }

                return header;
            }

            private static byte[] BuildMiniStreamHeader() {
                byte[] header = new byte[SectorSize];
                byte[] signature = { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 };
                Buffer.BlockCopy(signature, 0, header, 0, signature.Length);
                WriteUInt16(header, 24, 0x003e);
                WriteUInt16(header, 26, 0x0003);
                WriteUInt16(header, 28, 0xfffe);
                WriteUInt16(header, 30, 0x0009);
                WriteUInt16(header, 32, 0x0006);
                WriteUInt32(header, 44, 1);
                WriteUInt32(header, 48, 1);
                WriteUInt32(header, 56, 4096);
                WriteUInt32(header, 60, 2);
                WriteUInt32(header, 64, 1);
                WriteUInt32(header, 68, EndOfChain);
                for (int i = 0; i < 109; i++) {
                    WriteUInt32(header, 76 + i * 4, i == 0 ? 3U : FreeSect);
                }

                return header;
            }

            private static byte[] BuildDirectory(int workbookSize, bool includeVbaProjectStorage, bool includeOleObjectStorage) {
                byte[] directory = new byte[SectorSize];
                WriteDirectoryEntry(directory, 0, "Root Entry", 5, EndOfChain, EndOfChain, 1, EndOfChain, 0);
                WriteDirectoryEntry(directory, 128, "Workbook", 2, EndOfChain, includeVbaProjectStorage || includeOleObjectStorage ? 2 : EndOfChain, EndOfChain, 0, (ulong)workbookSize);
                if (includeVbaProjectStorage) {
                    WriteDirectoryEntry(directory, 256, "_VBA_PROJECT_CUR", 1, EndOfChain, includeOleObjectStorage ? 3U : EndOfChain, EndOfChain, EndOfChain, 0);
                }

                if (includeOleObjectStorage) {
                    WriteDirectoryEntry(directory, includeVbaProjectStorage ? 384 : 256, "ObjectPool", 1, EndOfChain, EndOfChain, EndOfChain, EndOfChain, 0);
                }

                return directory;
            }

            private static byte[] BuildRootStreamDirectory(IReadOnlyList<CompoundStreamSpec> streams) {
                byte[] directory = new byte[SectorSize];
                WriteDirectoryEntry(directory, 0, "Root Entry", 5, EndOfChain, EndOfChain, 1, EndOfChain, 0);

                uint startSector = 0;
                for (int i = 0; i < streams.Count; i++) {
                    CompoundStreamSpec stream = streams[i];
                    uint rightSibling = i + 1 < streams.Count ? (uint)(i + 2) : EndOfChain;
                    WriteDirectoryEntry(
                        directory,
                        (i + 1) * 128,
                        stream.Name,
                        2,
                        EndOfChain,
                        rightSibling,
                        EndOfChain,
                        startSector,
                        (ulong)stream.OriginalSize);
                    startSector += (uint)(stream.Bytes.Length / SectorSize);
                }

                return directory;
            }

            private static byte[] BuildNonExcelDirectory(int streamSize) {
                byte[] directory = new byte[SectorSize];
                WriteDirectoryEntry(directory, 0, "Root Entry", 5, EndOfChain, EndOfChain, 1, EndOfChain, 0);
                WriteDirectoryEntry(directory, 128, "NotExcel", 2, EndOfChain, EndOfChain, EndOfChain, 0, (ulong)streamSize);
                return directory;
            }

            private static byte[] BuildDifatDirectory(int workbookSize) {
                byte[] directory = new byte[SectorSize];
                WriteDirectoryEntry(directory, 0, "Root Entry", 5, EndOfChain, EndOfChain, 1, EndOfChain, 0);
                WriteDirectoryEntry(directory, 128, "Workbook", 2, EndOfChain, EndOfChain, EndOfChain, 0, (ulong)workbookSize);
                return directory;
            }

            private static byte[] BuildMiniStreamDirectory(int workbookSize, int rootMiniStreamSize) {
                byte[] directory = new byte[SectorSize];
                WriteDirectoryEntry(directory, 0, "Root Entry", 5, EndOfChain, EndOfChain, 1, 0, (ulong)rootMiniStreamSize);
                WriteDirectoryEntry(directory, 128, "Workbook", 2, EndOfChain, EndOfChain, EndOfChain, 0, (ulong)workbookSize);
                return directory;
            }

            private static byte[] BuildFat(int workbookSectorCount) {
                byte[] fat = new byte[SectorSize];
                for (int i = 0; i < workbookSectorCount; i++) {
                    WriteUInt32(fat, i * 4, i + 1 == workbookSectorCount ? EndOfChain : (uint)(i + 1));
                }

                WriteUInt32(fat, workbookSectorCount * 4, EndOfChain);
                WriteUInt32(fat, (workbookSectorCount + 1) * 4, FatSect);
                for (int offset = (workbookSectorCount + 2) * 4; offset < fat.Length; offset += 4) {
                    WriteUInt32(fat, offset, FreeSect);
                }

                return fat;
            }

            private static byte[] BuildRootStreamFat(IReadOnlyList<CompoundStreamSpec> streams) {
                int dataSectorCount = streams.Sum(stream => stream.Bytes.Length / SectorSize);
                byte[] fat = new byte[SectorSize];
                int sector = 0;
                foreach (CompoundStreamSpec stream in streams) {
                    int streamSectorCount = stream.Bytes.Length / SectorSize;
                    for (int i = 0; i < streamSectorCount; i++) {
                        WriteUInt32(fat, (sector + i) * 4, i + 1 == streamSectorCount ? EndOfChain : (uint)(sector + i + 1));
                    }

                    sector += streamSectorCount;
                }

                WriteUInt32(fat, dataSectorCount * 4, EndOfChain);
                WriteUInt32(fat, (dataSectorCount + 1) * 4, FatSect);
                for (int offset = (dataSectorCount + 2) * 4; offset < fat.Length; offset += 4) {
                    WriteUInt32(fat, offset, FreeSect);
                }

                return fat;
            }

            private static byte[] BuildDifatFat(int workbookSectorCount, int directorySector, int firstFatSector, int fatSectorCount, int difatSector) {
                byte[] fat = new byte[fatSectorCount * SectorSize];
                for (int offset = 0; offset < fat.Length; offset += 4) {
                    WriteUInt32(fat, offset, FreeSect);
                }

                for (int i = 0; i < workbookSectorCount; i++) {
                    WriteFatEntry(fat, i, i + 1 == workbookSectorCount ? EndOfChain : (uint)(i + 1));
                }

                WriteFatEntry(fat, directorySector, EndOfChain);
                for (int i = 0; i < fatSectorCount; i++) {
                    WriteFatEntry(fat, firstFatSector + i, FatSect);
                }

                WriteFatEntry(fat, difatSector, DifSect);
                return fat;
            }

            private static byte[] BuildDifatSector(int fatSectorFromDifat) {
                byte[] difat = new byte[SectorSize];
                for (int offset = 0; offset < difat.Length; offset += 4) {
                    WriteUInt32(difat, offset, FreeSect);
                }

                WriteUInt32(difat, 0, (uint)fatSectorFromDifat);
                WriteUInt32(difat, SectorSize - 4, EndOfChain);
                return difat;
            }

            private static byte[] BuildMiniFat(int usedMiniSectorCount) {
                byte[] miniFat = new byte[SectorSize];
                for (int i = 0; i < usedMiniSectorCount; i++) {
                    WriteUInt32(miniFat, i * 4, i + 1 == usedMiniSectorCount ? EndOfChain : (uint)(i + 1));
                }

                for (int offset = usedMiniSectorCount * 4; offset < miniFat.Length; offset += 4) {
                    WriteUInt32(miniFat, offset, FreeSect);
                }

                return miniFat;
            }

            private static byte[] BuildMiniStreamFat() {
                byte[] fat = new byte[SectorSize];
                WriteUInt32(fat, 0, EndOfChain);
                WriteUInt32(fat, 4, EndOfChain);
                WriteUInt32(fat, 8, EndOfChain);
                WriteUInt32(fat, 12, FatSect);
                for (int offset = 16; offset < fat.Length; offset += 4) {
                    WriteUInt32(fat, offset, FreeSect);
                }

                return fat;
            }

            private static void WriteDirectoryEntry(byte[] buffer, int offset, string name, byte objectType, uint left, uint right, uint child, uint startSector, ulong size) {
                byte[] nameBytes = Encoding.Unicode.GetBytes(name + '\0');
                Buffer.BlockCopy(nameBytes, 0, buffer, offset, nameBytes.Length);
                WriteUInt16(buffer, offset + 64, (ushort)nameBytes.Length);
                buffer[offset + 66] = objectType;
                buffer[offset + 67] = 1;
                WriteUInt32(buffer, offset + 68, left);
                WriteUInt32(buffer, offset + 72, right);
                WriteUInt32(buffer, offset + 76, child);
                WriteUInt32(buffer, offset + 116, startSector);
                WriteUInt64(buffer, offset + 120, size);
            }

            private static byte[] PadToRegularStream(byte[] bytes) {
                int regularStreamLength = Math.Max(4096, ((bytes.Length + SectorSize - 1) / SectorSize) * SectorSize);
                byte[] padded = new byte[regularStreamLength];
                Buffer.BlockCopy(bytes, 0, padded, 0, bytes.Length);
                return padded;
            }

            private static byte[] PadToSectorCount(byte[] bytes, int sectorCount) {
                byte[] padded = new byte[checked(sectorCount * SectorSize)];
                Buffer.BlockCopy(bytes, 0, padded, 0, bytes.Length);
                return padded;
            }

            private static byte[] PadToMiniStreamContainer(byte[] bytes) {
                int miniStreamLength = Math.Max(MiniSectorSize, ((bytes.Length + MiniSectorSize - 1) / MiniSectorSize) * MiniSectorSize);
                int containerLength = ((miniStreamLength + SectorSize - 1) / SectorSize) * SectorSize;
                byte[] padded = new byte[containerLength];
                Buffer.BlockCopy(bytes, 0, padded, 0, bytes.Length);
                return padded;
            }

            private static void WriteFatEntry(byte[] fat, int sector, uint value) {
                WriteUInt32(fat, checked(sector * 4), value);
            }

            private sealed class CompoundStreamSpec {
                internal CompoundStreamSpec(string name, byte[] bytes) : this(name, bytes, bytes.Length) {
                }

                internal CompoundStreamSpec(string name, byte[] bytes, int originalSize) {
                    Name = name;
                    Bytes = bytes;
                    OriginalSize = originalSize;
                }

                internal string Name { get; }

                internal byte[] Bytes { get; }

                internal int OriginalSize { get; }
            }
        }
    }
}
