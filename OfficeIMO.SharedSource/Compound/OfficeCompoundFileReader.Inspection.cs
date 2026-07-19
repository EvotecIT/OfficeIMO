using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

#if OFFICEIMO_READER_CORE
namespace OfficeIMO.Reader.Internal.Compound {
#else
namespace OfficeIMO.Drawing.Internal {
#endif
    internal static partial class OfficeCompoundFileReader {
        /// <summary>
        /// Inspects only the allocation tables and directory stream of a seekable compound file.
        /// Payload streams are never materialized.
        /// </summary>
        internal static bool TryContainsStreamPath(Stream stream, string expectedPath, long maxInputBytes,
            int maxDirectoryEntries, out bool contains, out string? error) {
            return TryContainsStreamPath(stream, expectedPath,
                maxInputBytes, maxDirectoryEntries, CancellationToken.None,
                out contains, out error);
        }

        internal static bool TryContainsStreamPath(Stream stream,
            string expectedPath, long maxInputBytes,
            int maxDirectoryEntries, CancellationToken cancellationToken,
            out bool contains, out string? error) {
            contains = false;
            if (string.IsNullOrWhiteSpace(expectedPath)) {
                throw new ArgumentException("A stream path is required.", nameof(expectedPath));
            }

            if (!TryInspectDirectory(stream, maxInputBytes, maxDirectoryEntries,
                cancellationToken,
                out IReadOnlyList<OfficeCompoundFileEntry> entries,
                out error)) {
                return false;
            }

            foreach (OfficeCompoundFileEntry entry in entries) {
                cancellationToken.ThrowIfCancellationRequested();
                if (entry.IsStream && string.Equals(entry.Path, expectedPath,
                        StringComparison.OrdinalIgnoreCase)) {
                    contains = true;
                    break;
                }
            }
            return true;
        }

        /// <summary>
        /// Inspects allocation tables and directory metadata without materializing payload streams.
        /// </summary>
        internal static bool TryInspectDirectory(Stream stream, long maxInputBytes,
            int maxDirectoryEntries, out IReadOnlyList<OfficeCompoundFileEntry> entries, out string? error) {
            return TryInspectDirectory(stream, maxInputBytes,
                maxDirectoryEntries, CancellationToken.None, out entries,
                out error);
        }

        internal static bool TryInspectDirectory(Stream stream,
            long maxInputBytes, int maxDirectoryEntries,
            CancellationToken cancellationToken,
            out IReadOnlyList<OfficeCompoundFileEntry> entries,
            out string? error) {
            entries = Array.Empty<OfficeCompoundFileEntry>();
            error = null;
            try {
                if (stream == null) throw new ArgumentNullException(nameof(stream));
                cancellationToken.ThrowIfCancellationRequested();
                if (!stream.CanRead || !stream.CanSeek) {
                    error = "Compound directory inspection requires a readable seekable stream.";
                    return false;
                }
                if (maxInputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxInputBytes));
                if (maxDirectoryEntries <= 0) throw new ArgumentOutOfRangeException(nameof(maxDirectoryEntries));

                long basePosition = stream.Position;
                long remainingBytes = checked(stream.Length - basePosition);
                if (remainingBytes < HeaderSize || remainingBytes > maxInputBytes) {
                    error = remainingBytes > maxInputBytes
                        ? $"The compound file exceeds the configured input bound of {maxInputBytes}."
                        : "The compound file is shorter than its header.";
                    return false;
                }

                byte[] header = ReadAt(stream, basePosition, HeaderSize,
                    cancellationToken);
                cancellationToken.ThrowIfCancellationRequested();
                if (!HasSignature(header)) {
                    error = "The file does not start with the OLE compound document signature.";
                    return false;
                }

                ushort majorVersion = ReadUInt16(header, 26);
                ushort byteOrder = ReadUInt16(header, 28);
                ushort sectorShift = ReadUInt16(header, 30);
                ushort miniSectorShift = ReadUInt16(header, 32);
                if ((sectorShift != 9 && sectorShift != 12) || miniSectorShift != 6) {
                    error = "The compound file has unsupported sector sizes.";
                    return false;
                }
                int sectorSize = 1 << sectorShift;
                bool validVersion = (majorVersion == 3 && sectorSize == 512) ||
                    (majorVersion == 4 && sectorSize == 4096);
                if (!validVersion || byteOrder != 0xfffe || remainingBytes < sectorSize) {
                    error = "The compound file has an unsupported version or byte order.";
                    return false;
                }

                int physicalSectorCount = checked((int)((remainingBytes - sectorSize) / sectorSize));
                int fatSectorCount = checked((int)ReadUInt32(header, 44));
                uint directoryStart = ReadUInt32(header, 48);
                uint firstDifat = ReadUInt32(header, 68);
                int difatSectorCount = checked((int)ReadUInt32(header, 72));
                if (fatSectorCount < 0 || fatSectorCount > physicalSectorCount || difatSectorCount < 0 ||
                    difatSectorCount > physicalSectorCount) {
                    throw new InvalidDataException("Compound allocation table counts exceed the file size.");
                }

                List<uint> fatSectorIds = ReadFatSectorIds(stream, basePosition, header, sectorSize,
                    physicalSectorCount, firstDifat, difatSectorCount,
                    fatSectorCount, cancellationToken);
                byte[] directoryBytes = ReadDirectoryStream(stream, basePosition, directoryStart, sectorSize,
                    physicalSectorCount, fatSectorIds, maxDirectoryEntries,
                    cancellationToken);
                List<DirectoryEntry> directoryEntries = ReadDirectoryEntries(directoryBytes, majorVersion,
                    maxDirectoryEntries, cancellationToken);
                entries = BuildCompoundEntries(directoryEntries,
                    cancellationToken);
                cancellationToken.ThrowIfCancellationRequested();
                return true;
            } catch (Exception exception) when (exception is IOException || exception is ArgumentException ||
                exception is InvalidDataException || exception is OverflowException ||
                exception is IndexOutOfRangeException || exception is NotSupportedException) {
                entries = Array.Empty<OfficeCompoundFileEntry>();
                error = $"The OLE compound directory could not be inspected. {exception.Message}";
                return false;
            }
        }

        private static List<uint> ReadFatSectorIds(Stream stream, long basePosition, byte[] header,
            int sectorSize, int physicalSectorCount, uint firstDifat,
            int difatSectorCount, int fatSectorCount,
            CancellationToken cancellationToken) {
            var result = new List<uint>(fatSectorCount);
            var visitedFatSectors = new HashSet<uint>();
            for (int index = 0; index < 109 && result.Count < fatSectorCount; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                AddFatSector(ReadUInt32(header, 76 + index * 4), physicalSectorCount, visitedFatSectors, result);
            }

            uint nextDifat = firstDifat;
            var visitedDifatSectors = new HashSet<uint>();
            int entriesPerDifatSector = sectorSize / 4 - 1;
            for (int index = 0; index < difatSectorCount && nextDifat != EndOfChain &&
                result.Count < fatSectorCount; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (nextDifat >= physicalSectorCount || !visitedDifatSectors.Add(nextDifat)) {
                    throw new InvalidDataException("Invalid compound DIFAT sector chain.");
                }
                byte[] difat = ReadSector(stream, basePosition, nextDifat,
                    sectorSize, physicalSectorCount, cancellationToken);
                for (int entry = 0; entry < entriesPerDifatSector && result.Count < fatSectorCount; entry++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    AddFatSector(ReadUInt32(difat, entry * 4), physicalSectorCount, visitedFatSectors, result);
                }
                nextDifat = ReadUInt32(difat, entriesPerDifatSector * 4);
            }

            if (result.Count != fatSectorCount) {
                throw new InvalidDataException("The DIFAT does not reference the declared FAT sector count.");
            }
            return result;
        }

        private static void AddFatSector(uint sector, int physicalSectorCount, ISet<uint> visited,
            ICollection<uint> result) {
            if (sector == FreeSect) return;
            if (sector >= physicalSectorCount || !visited.Add(sector)) {
                throw new InvalidDataException("Invalid or duplicate FAT sector reference.");
            }
            result.Add(sector);
        }

        private static byte[] ReadDirectoryStream(Stream stream, long basePosition, uint directoryStart,
            int sectorSize, int physicalSectorCount, IReadOnlyList<uint> fatSectorIds,
            int maxDirectoryEntries, CancellationToken cancellationToken) {
            long maximumBytes = checked((long)maxDirectoryEntries * DirectoryEntrySize);
            using (var output = new MemoryStream()) {
                uint sector = directoryStart;
                var visited = new HashSet<uint>();
                var fatCache = new Dictionary<uint, byte[]>();
                while (sector != EndOfChain && sector != FreeSect) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (sector >= physicalSectorCount || !visited.Add(sector)) {
                        throw new InvalidDataException("Invalid compound directory sector chain.");
                    }
                    if (output.Length > maximumBytes - sectorSize) {
                        throw new InvalidDataException($"Compound directory entry count exceeds {maxDirectoryEntries}.");
                    }
                    byte[] directorySector = ReadSector(stream, basePosition, sector, sectorSize,
                        physicalSectorCount, cancellationToken);
                    output.Write(directorySector, 0, directorySector.Length);
                    sector = ReadFatEntry(stream, basePosition, sector, sectorSize, physicalSectorCount,
                        fatSectorIds, fatCache, cancellationToken);
                }
                return output.ToArray();
            }
        }

        private static uint ReadFatEntry(Stream stream, long basePosition, uint sector, int sectorSize,
            int physicalSectorCount, IReadOnlyList<uint> fatSectorIds,
            IDictionary<uint, byte[]> cache,
            CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            int entriesPerSector = sectorSize / 4;
            int fatSectorIndex = checked((int)(sector / entriesPerSector));
            if (fatSectorIndex >= fatSectorIds.Count) {
                throw new InvalidDataException("The FAT does not contain the requested sector entry.");
            }
            uint fatSectorId = fatSectorIds[fatSectorIndex];
            if (!cache.TryGetValue(fatSectorId, out byte[]? fatSector)) {
                fatSector = ReadSector(stream, basePosition, fatSectorId,
                    sectorSize, physicalSectorCount, cancellationToken);
                cache[fatSectorId] = fatSector;
            }
            return ReadUInt32(fatSector, checked((int)(sector % entriesPerSector)) * 4);
        }

        private static byte[] ReadSector(Stream stream, long basePosition, uint sector, int sectorSize,
            int physicalSectorCount,
            CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            if (sector >= physicalSectorCount) throw new InvalidDataException("Compound sector points outside the file.");
            long offset = checked(basePosition + checked(((long)sector + 1) * sectorSize));
            return ReadAt(stream, offset, sectorSize, cancellationToken);
        }

        private static byte[] ReadAt(Stream stream, long offset, int count,
            CancellationToken cancellationToken) {
            byte[] buffer = new byte[count];
            stream.Position = offset;
            int total = 0;
            while (total < count) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = stream.Read(buffer, total, count - total);
                if (read <= 0) throw new EndOfStreamException("The compound file ended before the requested metadata was read.");
                total += read;
            }
            cancellationToken.ThrowIfCancellationRequested();
            return buffer;
        }
    }
}
