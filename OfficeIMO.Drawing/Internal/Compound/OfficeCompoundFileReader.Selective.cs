using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Drawing.Internal {
    internal static partial class OfficeCompoundFileReader {
        /// <summary>Reads metadata streams while copying selected payload streams to caller-owned destinations.</summary>
        internal static bool TryReadSelective(Stream stream, OfficeCompoundReadOptions options,
            Func<string, long, bool> externalize,
            Func<string, long, Stream> openExternalDestination,
            out OfficeCompoundFile? compoundFile, out string? error) =>
            TryReadSelective(stream, options, externalize,
                openExternalDestination, CancellationToken.None,
                out compoundFile, out error);

        internal static bool TryReadSelective(Stream stream,
            OfficeCompoundReadOptions options,
            Func<string, long, bool> externalize,
            Func<string, long, Stream> openExternalDestination,
            CancellationToken cancellationToken,
            out OfficeCompoundFile? compoundFile, out string? error) {
            compoundFile = null;
            error = null;
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (externalize == null) throw new ArgumentNullException(nameof(externalize));
            if (openExternalDestination == null) throw new ArgumentNullException(nameof(openExternalDestination));
            if (!stream.CanRead || !stream.CanSeek) {
                error = "Selective compound reading requires a readable seekable stream.";
                return false;
            }

            long originalPosition = stream.Position;
            try {
                cancellationToken.ThrowIfCancellationRequested();
                long basePosition = originalPosition;
                long remainingBytes = checked(stream.Length - basePosition);
                if (remainingBytes < HeaderSize) {
                    error = "The compound file is shorter than its header.";
                    return false;
                }
                byte[] header = ReadAt(stream, basePosition, HeaderSize,
                    cancellationToken);
                if (!HasSignature(header)) {
                    error = "The file does not start with the OLE compound document signature.";
                    return false;
                }

                ushort majorVersion = ReadUInt16(header, 26);
                ushort byteOrder = ReadUInt16(header, 28);
                ushort sectorShift = ReadUInt16(header, 30);
                ushort miniSectorShift = ReadUInt16(header, 32);
                if ((sectorShift != 9 && sectorShift != 12) || miniSectorShift != 6) {
                    throw new InvalidDataException("Unsupported compound file sector sizes.");
                }
                int sectorSize = 1 << sectorShift;
                bool validVersion = (majorVersion == 3 && sectorSize == 512) ||
                    (majorVersion == 4 && sectorSize == 4096);
                if (!validVersion || byteOrder != 0xfffe || remainingBytes < sectorSize) {
                    throw new InvalidDataException("Unsupported compound file version or byte order.");
                }

                int physicalSectorCount = checked((int)((remainingBytes - sectorSize) / sectorSize));
                int fatSectorCount = checked((int)ReadUInt32(header, 44));
                uint directoryStart = ReadUInt32(header, 48);
                uint miniCutoff = ReadUInt32(header, 56);
                uint miniFatStart = ReadUInt32(header, 60);
                int miniFatSectorCount = checked((int)ReadUInt32(header, 64));
                uint firstDifat = ReadUInt32(header, 68);
                int difatSectorCount = checked((int)ReadUInt32(header, 72));
                if (fatSectorCount > physicalSectorCount || difatSectorCount > physicalSectorCount) {
                    throw new InvalidDataException("Compound allocation table counts exceed the file size.");
                }

                List<uint> fatSectorIds = ReadFatSectorIds(stream, basePosition, header, sectorSize,
                    physicalSectorCount, firstDifat, difatSectorCount,
                    fatSectorCount, cancellationToken);
                byte[] directoryBytes = ReadDirectoryStream(stream, basePosition, directoryStart, sectorSize,
                    physicalSectorCount, fatSectorIds,
                    options.MaxDirectoryEntries, cancellationToken);
                List<DirectoryEntry> entries = ReadDirectoryEntries(directoryBytes, majorVersion,
                    options.MaxDirectoryEntries);
                DirectoryEntry? root = entries.FirstOrDefault(entry => entry.ObjectType == 5);
                if (root == null) throw new InvalidDataException("Compound file root directory entry is missing.");
                if (root.Size < 0 || root.Size > options.MaxTotalStreamBytes || root.Size > remainingBytes) {
                    throw new InvalidDataException("Compound file mini stream exceeds configured or physical bounds.");
                }

                IReadOnlyDictionary<int, string> streamPaths = BuildCompoundEntryPaths(entries);
                DirectoryEntry[] streamEntries = entries.Where(entry => entry.ObjectType == 2).ToArray();
                if (streamEntries.Length > options.MaxStreamCount) {
                    throw new InvalidDataException($"Compound stream count {streamEntries.Length} exceeds {options.MaxStreamCount}.");
                }
                long totalStreamBytes = 0;
                var externalStreams = new HashSet<int>();
                foreach (DirectoryEntry entry in streamEntries) {
                    string path = streamPaths.TryGetValue(entry.Index, out string? entryPath) ? entryPath : entry.Name;
                    bool isExternal = externalize(path, entry.Size);
                    if (isExternal) externalStreams.Add(entry.Index);
                    if (entry.Size < 0 || entry.Size > options.MaxStreamBytes ||
                        (!isExternal && entry.Size > int.MaxValue)) {
                        throw new InvalidDataException($"Compound stream '{path}' has unsupported size {entry.Size}.");
                    }
                    totalStreamBytes = checked(totalStreamBytes + entry.Size);
                    if (totalStreamBytes > options.MaxTotalStreamBytes) {
                        throw new InvalidDataException($"Compound stream bytes exceed {options.MaxTotalStreamBytes}.");
                    }
                    options.StreamSizeValidator?.Invoke(path, entry.Size);
                }

                var fatCache = new Dictionary<uint, byte[]>();
                uint[] miniFat = miniFatStart == EndOfChain || miniFatSectorCount == 0
                    ? Array.Empty<uint>()
                    : BytesToUInt32Array(ReadRegularChain(stream, basePosition, miniFatStart,
                        checked((long)miniFatSectorCount * sectorSize), sectorSize, physicalSectorCount,
                        fatSectorIds, fatCache, cancellationToken));
                List<uint> rootChain = root.StartSector == EndOfChain || root.Size == 0
                    ? new List<uint>()
                    : GetRegularSectorChain(stream, basePosition, root.StartSector, root.Size, sectorSize,
                        physicalSectorCount, fatSectorIds, fatCache,
                        cancellationToken);

                var streams = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
                foreach (DirectoryEntry entry in streamEntries) {
                    string path = streamPaths.TryGetValue(entry.Index, out string? entryPath) ? entryPath : entry.Name;
                    if (externalStreams.Contains(entry.Index)) {
                        using (Stream destination = openExternalDestination(path, entry.Size)) {
                            if (destination == null || !destination.CanWrite) {
                                throw new InvalidDataException("The external compound destination is not writable.");
                            }
                            CopyEntry(stream, destination, entry, miniCutoff, miniFat, rootChain, basePosition,
                                sectorSize, physicalSectorCount, fatSectorIds,
                                fatCache, cancellationToken);
                        }
                        // Retain the property path so higher-level parsers can recognize an externally-backed value.
                        streams[path] = Array.Empty<byte>();
                    } else {
                        using (var destination = new MemoryStream(checked((int)entry.Size))) {
                            CopyEntry(stream, destination, entry, miniCutoff, miniFat, rootChain, basePosition,
                                sectorSize, physicalSectorCount, fatSectorIds,
                                fatCache, cancellationToken);
                            streams[path] = destination.ToArray();
                        }
                    }
                    if (string.Equals(path, entry.Name, StringComparison.OrdinalIgnoreCase)) {
                        streams[entry.Name] = streams[path];
                    }
                }

                compoundFile = new OfficeCompoundFile(streams, BuildCompoundEntries(entries),
                    CreateCompoundEntry(root, "Root Entry"));
                return true;
            } catch (Exception exception) when (exception is IOException || exception is ArgumentException ||
                exception is InvalidDataException || exception is OverflowException ||
                exception is IndexOutOfRangeException || exception is NotSupportedException) {
                compoundFile = null;
                error = $"The OLE compound file could not be read selectively. {exception.Message}";
                return false;
            } finally {
                stream.Position = originalPosition;
            }
        }

        private static void CopyEntry(Stream input, Stream output, DirectoryEntry entry, uint miniCutoff,
            IReadOnlyList<uint> miniFat, IReadOnlyList<uint> rootChain, long basePosition, int sectorSize,
            int physicalSectorCount, IReadOnlyList<uint> fatSectorIds,
            IDictionary<uint, byte[]> fatCache,
            CancellationToken cancellationToken) {
            if (entry.Size == 0) return;
            if (entry.Size < miniCutoff) {
                CopyMiniChain(input, output, entry.StartSector, entry.Size, miniFat, rootChain,
                    basePosition, sectorSize, cancellationToken);
                return;
            }
            CopyRegularChain(input, output, entry.StartSector, entry.Size, basePosition, sectorSize,
                physicalSectorCount, fatSectorIds, fatCache,
                cancellationToken);
        }

        private static byte[] ReadRegularChain(Stream input, long basePosition, uint startSector, long size,
            int sectorSize, int physicalSectorCount, IReadOnlyList<uint> fatSectorIds,
            IDictionary<uint, byte[]> fatCache,
            CancellationToken cancellationToken) {
            using (var output = new MemoryStream(checked((int)size))) {
                CopyRegularChain(input, output, startSector, size, basePosition, sectorSize,
                    physicalSectorCount, fatSectorIds, fatCache,
                    cancellationToken);
                return output.ToArray();
            }
        }

        private static void CopyRegularChain(Stream input, Stream output, uint startSector, long size,
            long basePosition, int sectorSize, int physicalSectorCount, IReadOnlyList<uint> fatSectorIds,
            IDictionary<uint, byte[]> fatCache,
            CancellationToken cancellationToken) {
            uint sector = startSector;
            long remaining = size;
            var visited = new HashSet<uint>();
            while (remaining > 0) {
                if (sector == EndOfChain || sector == FreeSect || sector >= physicalSectorCount ||
                    !visited.Add(sector)) {
                    throw new InvalidDataException("Compound sector chain is shorter than its declared stream size.");
                }
                byte[] bytes = ReadSector(input, basePosition, sector,
                    sectorSize, physicalSectorCount, cancellationToken);
                int write = (int)Math.Min(bytes.Length, remaining);
                output.Write(bytes, 0, write);
                remaining -= write;
                sector = ReadFatEntry(input, basePosition, sector, sectorSize, physicalSectorCount,
                    fatSectorIds, fatCache, cancellationToken);
            }
        }

        private static List<uint> GetRegularSectorChain(Stream input, long basePosition, uint startSector,
            long size, int sectorSize, int physicalSectorCount, IReadOnlyList<uint> fatSectorIds,
            IDictionary<uint, byte[]> fatCache,
            CancellationToken cancellationToken) {
            int required = checked((int)((size + sectorSize - 1) / sectorSize));
            var result = new List<uint>(required);
            var visited = new HashSet<uint>();
            uint sector = startSector;
            while (result.Count < required) {
                if (sector == EndOfChain || sector == FreeSect || sector >= physicalSectorCount ||
                    !visited.Add(sector)) {
                    throw new InvalidDataException("Compound mini stream chain is shorter than its declared size.");
                }
                result.Add(sector);
                sector = ReadFatEntry(input, basePosition, sector, sectorSize, physicalSectorCount,
                    fatSectorIds, fatCache, cancellationToken);
            }
            return result;
        }

        private static void CopyMiniChain(Stream input, Stream output, uint startSector, long size,
            IReadOnlyList<uint> miniFat, IReadOnlyList<uint> rootChain,
            long basePosition, int sectorSize,
            CancellationToken cancellationToken) {
            uint miniSector = startSector;
            long remaining = size;
            var visited = new HashSet<uint>();
            while (remaining > 0) {
                if (miniSector == EndOfChain || miniSector == FreeSect || miniSector >= miniFat.Count ||
                    !visited.Add(miniSector)) {
                    throw new InvalidDataException("Compound mini-sector chain is shorter than its declared size.");
                }
                long miniOffset = checked((long)miniSector * MiniSectorSize);
                int rootSectorIndex = checked((int)(miniOffset / sectorSize));
                int offsetWithinSector = checked((int)(miniOffset % sectorSize));
                if (rootSectorIndex >= rootChain.Count || offsetWithinSector + MiniSectorSize > sectorSize) {
                    throw new InvalidDataException("Compound mini-sector points outside the root mini stream.");
                }
                long physicalOffset = checked(basePosition + ((long)rootChain[rootSectorIndex] + 1) * sectorSize +
                    offsetWithinSector);
                byte[] bytes = ReadAt(input, physicalOffset,
                    MiniSectorSize, cancellationToken);
                int write = (int)Math.Min(bytes.Length, remaining);
                output.Write(bytes, 0, write);
                remaining -= write;
                miniSector = miniFat[(int)miniSector];
            }
        }
    }
}
