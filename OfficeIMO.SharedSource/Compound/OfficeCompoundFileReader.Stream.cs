using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Drawing.Internal {
    internal static partial class OfficeCompoundFileReader {
        /// <summary>
        /// Opens one compound payload as a bounded seekable view without materializing its bytes.
        /// The returned stream owns <paramref name="source"/> unless <paramref name="leaveOpen"/> is true.
        /// </summary>
        internal static bool TryOpenStream(
            Stream source,
            OfficeCompoundReadOptions options,
            Func<string, long, bool> selector,
            bool leaveOpen,
            CancellationToken cancellationToken,
            out Stream? selectedStream,
            out string? error) {
            selectedStream = null;
            error = null;
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (selector == null) throw new ArgumentNullException(nameof(selector));
            if (!source.CanRead || !source.CanSeek) {
                error = "Compound stream views require a readable seekable source.";
                return false;
            }

            long originalPosition = source.Position;
            try {
                cancellationToken.ThrowIfCancellationRequested();
                long basePosition = originalPosition;
                long remainingBytes = checked(source.Length - basePosition);
                if (remainingBytes < HeaderSize) {
                    source.Position = originalPosition;
                    error = "The compound file is shorter than its header.";
                    return false;
                }
                byte[] header = ReadAt(source, basePosition, HeaderSize, cancellationToken);
                if (!HasSignature(header)) {
                    source.Position = originalPosition;
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
                bool validVersion = (majorVersion == 3 && sectorSize == 512)
                    || (majorVersion == 4 && sectorSize == 4096);
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
                ValidateAllocationTableCounts(
                    fatSectorCount,
                    difatSectorCount,
                    miniFatSectorCount,
                    physicalSectorCount,
                    sectorSize,
                    options.MaxTotalStreamBytes);

                List<uint> fatSectorIds = ReadFatSectorIds(
                    source,
                    basePosition,
                    header,
                    sectorSize,
                    physicalSectorCount,
                    firstDifat,
                    difatSectorCount,
                    fatSectorCount,
                    cancellationToken);
                byte[] directoryBytes = ReadDirectoryStream(
                    source,
                    basePosition,
                    directoryStart,
                    sectorSize,
                    physicalSectorCount,
                    fatSectorIds,
                    options.MaxDirectoryEntries,
                    cancellationToken);
                List<DirectoryEntry> entries = ReadDirectoryEntries(
                    directoryBytes,
                    majorVersion,
                    options.MaxDirectoryEntries,
                    cancellationToken);
                DirectoryEntry? root = entries.FirstOrDefault(static entry => entry.ObjectType == 5);
                if (root == null) throw new InvalidDataException("Compound file root directory entry is missing.");
                long maximumPhysicalStreamBytes = checked((long)physicalSectorCount * sectorSize);
                if (root.Size < 0
                    || root.Size > options.MaxTotalStreamBytes
                    || root.Size > maximumPhysicalStreamBytes) {
                    throw new InvalidDataException("Compound file mini stream exceeds configured or physical bounds.");
                }

                IReadOnlyDictionary<int, string> paths = BuildCompoundEntryPaths(entries, cancellationToken);
                DirectoryEntry? selected = null;
                foreach (DirectoryEntry entry in entries) {
                    if (entry.ObjectType != 2) continue;
                    cancellationToken.ThrowIfCancellationRequested();
                    string path = paths.TryGetValue(entry.Index, out string? value) ? value : entry.Name;
                    if (!selector(path, entry.Size)) continue;
                    if (selected != null) {
                        throw new InvalidDataException("More than one compound stream matched the requested selector.");
                    }
                    if (entry.Size < 0 || entry.Size > options.MaxStreamBytes) {
                        throw new InvalidDataException($"Compound stream '{path}' has unsupported size {entry.Size}.");
                    }
                    ValidateRegularStreamPhysicalBounds(path, entry.Size, miniCutoff, maximumPhysicalStreamBytes);
                    options.StreamSizeValidator?.Invoke(path, entry.Size);
                    selected = entry;
                }
                if (selected == null) {
                    source.Position = originalPosition;
                    error = "The requested compound stream was not found.";
                    return false;
                }

                List<uint>? regularChain = null;
                List<uint>? miniChain = null;
                List<uint>? rootChain = null;
                if (selected.Size >= miniCutoff) {
                    regularChain = GetRegularSectorChainForView(
                        source,
                        basePosition,
                        selected.StartSector,
                        selected.Size,
                        sectorSize,
                        physicalSectorCount,
                        fatSectorIds,
                        cancellationToken);
                } else if (selected.Size > 0) {
                    var fatCache = new Dictionary<uint, byte[]>();
                    uint[] miniFat = miniFatStart == EndOfChain || miniFatSectorCount == 0
                        ? Array.Empty<uint>()
                        : BytesToUInt32Array(ReadRegularChain(
                            source,
                            basePosition,
                            miniFatStart,
                            checked((long)miniFatSectorCount * sectorSize),
                            sectorSize,
                            physicalSectorCount,
                            fatSectorIds,
                            fatCache,
                            cancellationToken), cancellationToken);
                    rootChain = root.StartSector == EndOfChain || root.Size == 0
                        ? new List<uint>()
                        : GetRegularSectorChainForView(
                            source,
                            basePosition,
                            root.StartSector,
                            root.Size,
                            sectorSize,
                            physicalSectorCount,
                            fatSectorIds,
                            cancellationToken);
                    miniChain = GetMiniSectorChain(selected.StartSector, selected.Size, miniFat);
                }

                selectedStream = new CompoundEntryStream(
                    source,
                    basePosition,
                    sectorSize,
                    selected.Size,
                    regularChain,
                    miniChain,
                    rootChain,
                    leaveOpen,
                    cancellationToken);
                return true;
            } catch (Exception exception) when (exception is IOException
                                                || exception is ArgumentException
                                                || exception is InvalidDataException
                                                || exception is OverflowException
                                                || exception is IndexOutOfRangeException
                                                || exception is NotSupportedException) {
                source.Position = originalPosition;
                error = $"The OLE compound stream could not be opened. {exception.Message}";
                return false;
            }
        }

        private static List<uint> GetMiniSectorChain(uint startSector, long size, IReadOnlyList<uint> miniFat) {
            int required = checked((int)((size + MiniSectorSize - 1) / MiniSectorSize));
            var chain = new List<uint>(required);
            var visited = new bool[miniFat.Count];
            uint sector = startSector;
            while (chain.Count < required) {
                if (sector == EndOfChain || sector == FreeSect || sector >= miniFat.Count || visited[sector]) {
                    throw new InvalidDataException("Compound mini-sector chain is shorter than its declared stream size.");
                }
                visited[sector] = true;
                chain.Add(sector);
                sector = miniFat[checked((int)sector)];
            }
            return chain;
        }

        private static List<uint> GetRegularSectorChainForView(
            Stream source,
            long basePosition,
            uint startSector,
            long size,
            int sectorSize,
            int physicalSectorCount,
            IReadOnlyList<uint> fatSectorIds,
            CancellationToken cancellationToken) {
            int required = checked((int)((size + sectorSize - 1) / sectorSize));
            var chain = new List<uint>(required);
            var visited = new uint[checked((physicalSectorCount + 31) / 32)];
            byte[] fatSector = new byte[sectorSize];
            int loadedFatIndex = -1;
            int entriesPerSector = sectorSize / 4;
            uint sector = startSector;
            while (chain.Count < required) {
                if (sector == EndOfChain || sector == FreeSect || sector >= physicalSectorCount) {
                    throw new InvalidDataException("Compound sector chain is shorter than its declared stream size.");
                }
                int visitedIndex = checked((int)(sector / 32));
                uint visitedMask = 1U << checked((int)(sector % 32));
                if ((visited[visitedIndex] & visitedMask) != 0) {
                    throw new InvalidDataException("Compound sector chain contains a cycle.");
                }
                visited[visitedIndex] |= visitedMask;
                chain.Add(sector);

                int fatIndex = checked((int)(sector / entriesPerSector));
                if (fatIndex >= fatSectorIds.Count) {
                    throw new InvalidDataException("The FAT does not contain the requested sector entry.");
                }
                if (fatIndex != loadedFatIndex) {
                    long physical = checked(basePosition + ((long)fatSectorIds[fatIndex] + 1) * sectorSize);
                    ReadExact(source, physical, fatSector, cancellationToken);
                    loadedFatIndex = fatIndex;
                }
                sector = ReadUInt32(fatSector, checked((int)(sector % entriesPerSector)) * 4);
            }
            return chain;
        }

        private static void ReadExact(
            Stream source,
            long offset,
            byte[] buffer,
            CancellationToken cancellationToken) {
            source.Position = offset;
            int total = 0;
            while (total < buffer.Length) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = source.Read(buffer, total, buffer.Length - total);
                if (read <= 0) throw new EndOfStreamException("The compound file ended inside a FAT sector.");
                total += read;
            }
        }

        private sealed class CompoundEntryStream : Stream {
            private readonly Stream _source;
            private readonly long _basePosition;
            private readonly int _sectorSize;
            private readonly long _length;
            private readonly IReadOnlyList<uint>? _regularChain;
            private readonly IReadOnlyList<uint>? _miniChain;
            private readonly IReadOnlyList<uint>? _rootChain;
            private readonly bool _leaveOpen;
            private readonly CancellationToken _cancellationToken;
            private long _position;
            private bool _disposed;

            internal CompoundEntryStream(
                Stream source,
                long basePosition,
                int sectorSize,
                long length,
                IReadOnlyList<uint>? regularChain,
                IReadOnlyList<uint>? miniChain,
                IReadOnlyList<uint>? rootChain,
                bool leaveOpen,
                CancellationToken cancellationToken) {
                _source = source;
                _basePosition = basePosition;
                _sectorSize = sectorSize;
                _length = length;
                _regularChain = regularChain;
                _miniChain = miniChain;
                _rootChain = rootChain;
                _leaveOpen = leaveOpen;
                _cancellationToken = cancellationToken;
            }

            public override bool CanRead => !_disposed;
            public override bool CanSeek => !_disposed;
            public override bool CanWrite => false;
            public override long Length { get { ThrowIfDisposed(); return _length; } }
            public override long Position {
                get { ThrowIfDisposed(); return _position; }
                set { Seek(value, SeekOrigin.Begin); }
            }

            public override int Read(byte[] buffer, int offset, int count) {
                ThrowIfDisposed();
                if (buffer == null) throw new ArgumentNullException(nameof(buffer));
                if (offset < 0 || count < 0 || offset > buffer.Length - count) throw new ArgumentOutOfRangeException();
                if (_position >= _length || count == 0) return 0;
                _cancellationToken.ThrowIfCancellationRequested();
                int total = 0;
                int wanted = checked((int)Math.Min(count, _length - _position));
                while (total < wanted) {
                    int read = _regularChain != null
                        ? ReadRegular(buffer, offset + total, wanted - total)
                        : ReadMini(buffer, offset + total, wanted - total);
                    if (read <= 0) throw new EndOfStreamException("The compound stream ended before its declared length.");
                    total += read;
                    _position += read;
                }
                return total;
            }

            private int ReadRegular(byte[] buffer, int offset, int count) {
                int chainIndex = checked((int)(_position / _sectorSize));
                int within = checked((int)(_position % _sectorSize));
                int sectorCount = 1;
                if (within == 0) {
                    while (chainIndex + sectorCount < _regularChain!.Count
                           && _regularChain[chainIndex + sectorCount] == _regularChain[chainIndex] + sectorCount
                           && sectorCount * _sectorSize < count) {
                        sectorCount++;
                    }
                }
                int available = checked(sectorCount * _sectorSize - within);
                int read = Math.Min(count, available);
                long physical = checked(_basePosition + ((long)_regularChain![chainIndex] + 1) * _sectorSize + within);
                ReadSource(physical, buffer, offset, read);
                return read;
            }

            private int ReadMini(byte[] buffer, int offset, int count) {
                int chainIndex = checked((int)(_position / MiniSectorSize));
                int within = checked((int)(_position % MiniSectorSize));
                uint miniSector = _miniChain![chainIndex];
                long miniOffset = checked((long)miniSector * MiniSectorSize + within);
                int rootIndex = checked((int)(miniOffset / _sectorSize));
                int rootOffset = checked((int)(miniOffset % _sectorSize));
                if (rootIndex >= _rootChain!.Count) throw new InvalidDataException("Compound mini-sector points outside the root mini stream.");
                int read = Math.Min(count, MiniSectorSize - within);
                long physical = checked(_basePosition + ((long)_rootChain[rootIndex] + 1) * _sectorSize + rootOffset);
                ReadSource(physical, buffer, offset, read);
                return read;
            }

            private void ReadSource(long physical, byte[] buffer, int offset, int count) {
                _source.Position = physical;
                int total = 0;
                while (total < count) {
                    _cancellationToken.ThrowIfCancellationRequested();
                    int read = _source.Read(buffer, offset + total, count - total);
                    if (read <= 0) throw new EndOfStreamException("The compound source ended while reading a sector chain.");
                    total += read;
                }
            }

            public override long Seek(long offset, SeekOrigin origin) {
                ThrowIfDisposed();
                long position = origin switch {
                    SeekOrigin.Begin => offset,
                    SeekOrigin.Current => checked(_position + offset),
                    SeekOrigin.End => checked(_length + offset),
                    _ => throw new ArgumentOutOfRangeException(nameof(origin))
                };
                if (position < 0 || position > _length) throw new IOException("Attempted to seek outside the compound stream.");
                _position = position;
                return position;
            }

            public override void Flush() { }
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

            protected override void Dispose(bool disposing) {
                if (_disposed) return;
                _disposed = true;
                if (disposing && !_leaveOpen) _source.Dispose();
                base.Dispose(disposing);
            }

            private void ThrowIfDisposed() {
                if (_disposed) throw new ObjectDisposedException(nameof(CompoundEntryStream));
            }
        }
    }
}
