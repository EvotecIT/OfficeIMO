using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Core.Internal {
    internal static partial class OfficeCompoundFileReader {
        /// <summary>Indexes one immutable seekable compound source once. Views share the source cursor;
        /// this session and its views are for sequential use by one operation, not concurrent reads.</summary>
        internal sealed partial class ReadSession : IDisposable {
            private readonly Stream _source;
            private readonly bool _leaveOpen;
            private readonly CancellationToken _token;
            private readonly long _basePosition;
            private readonly int _sectorSize;
            private readonly int _physicalSectorCount;
            private readonly uint _miniCutoff;
            private readonly long _rootSize;
            private readonly uint _rootStartSector;
            private readonly List<uint> _fatSectorIds;
            private readonly Dictionary<uint, byte[]> _fatCache = new Dictionary<uint, byte[]>();
            private readonly Dictionary<string, DirectoryEntry> _streams = new Dictionary<string, DirectoryEntry>(StringComparer.Ordinal);
            private readonly uint[] _miniFat;
            private readonly List<uint> _rootChain;
            private Dictionary<string, List<uint>>? _validatedRegularChains;
            private Dictionary<string, List<uint>>? _validatedMiniChains;
            private bool _disposed;

            /// <summary>The source remains caller-owned on constructor failure. Ownership on success
            /// follows leaveOpen; disposing this session invalidates every outstanding view.</summary>
            internal ReadSession(Stream source, OfficeCompoundReadOptions options, bool leaveOpen, CancellationToken token,
                bool requireCanonicalHeader = false) {
                if (source == null) throw new ArgumentNullException(nameof(source));
                if (options == null) throw new ArgumentNullException(nameof(options));
                if (!source.CanRead || !source.CanSeek) throw new ArgumentException("Compound sessions require a readable seekable source.", nameof(source));
                _source = source; _leaveOpen = leaveOpen; _token = token; _basePosition = source.Position;
                try {
                    token.ThrowIfCancellationRequested();
                    long remaining = checked(source.Length - _basePosition);
                    if (remaining < HeaderSize || remaining > options.MaxTotalStreamBytes)
                        throw new InvalidDataException("The compound input is shorter than its header or exceeds the configured byte limit.");
                    byte[] header = ReadAt(source, _basePosition, HeaderSize, token);
                    if (!HasSignature(header)) throw new InvalidDataException("The compound signature is invalid.");
                    ushort major = ReadUInt16(header, 26);
                    ushort shift = ReadUInt16(header, 30);
                    if ((shift != 9 && shift != 12) || ReadUInt16(header, 32) != 6)
                        throw new InvalidDataException("Unsupported compound file sector sizes.");
                    _sectorSize = 1 << shift;
                    if (!((major == 3 && shift == 9) || (major == 4 && shift == 12)) || ReadUInt16(header, 28) != 0xfffe || remaining < _sectorSize)
                        throw new InvalidDataException("Unsupported compound file version or byte order.");
                    _physicalSectorCount = checked((int)((remaining - _sectorSize) / _sectorSize));
                    int fatCount = checked((int)ReadUInt32(header, 44));
                    uint miniFatStart = ReadUInt32(header, 60);
                    int miniFatCount = checked((int)ReadUInt32(header, 64));
                    uint difatStart = ReadUInt32(header, 68);
                    int difatCount = checked((int)ReadUInt32(header, 72));
                    _miniCutoff = ReadUInt32(header, 56);
                    if (requireCanonicalHeader && _miniCutoff != 4096)
                        throw new InvalidDataException("The compound mini-stream cutoff must be 4096 bytes.");
                    ValidateAllocationTableCounts(fatCount, difatCount, miniFatCount, _physicalSectorCount, _sectorSize, options.MaxTotalStreamBytes);
                    _fatSectorIds = ReadFatSectorIds(source, _basePosition, header, _sectorSize, _physicalSectorCount,
                        difatStart, difatCount, fatCount, token);
                    byte[] directory = ReadDirectoryStream(source, _basePosition, ReadUInt32(header, 48), _sectorSize,
                        _physicalSectorCount, _fatSectorIds, options.MaxDirectoryEntries, token);
                    List<DirectoryEntry> entries = ReadDirectoryEntries(directory, major, options.MaxDirectoryEntries, token);
                    DirectoryEntry? root = entries.FirstOrDefault(entry => entry.ObjectType == 5);
                    if (root == null) throw new InvalidDataException("Compound root directory entry is missing.");
                    RootEntry = CreateCompoundEntry(root, string.Empty);
                    HasCanonicalRootDirectory = root.Index == 0 && string.Equals(root.Name, "Root Entry", StringComparison.Ordinal)
                        && entries.Count(entry => entry.ObjectType == 5) == 1;
                    long physicalBytes = checked((long)_physicalSectorCount * _sectorSize);
                    if (root.Size < 0 || root.Size > options.MaxTotalStreamBytes || root.Size > physicalBytes)
                        throw new InvalidDataException("Compound mini stream exceeds configured or physical bounds.");
                    _rootSize = root.Size;
                    _rootStartSector = root.StartSector;
                    Entries = BuildCompoundEntries(entries, out bool hasInvalidDirectoryTree, token);
                    HasInvalidDirectoryTree = hasInvalidDirectoryTree;
                    IReadOnlyDictionary<int, string> paths = BuildCompoundEntryPaths(entries, token);
                    long total = 0;
                    foreach (DirectoryEntry entry in entries.Where(entry => entry.ObjectType == 2)) {
                        token.ThrowIfCancellationRequested();
                        string path = paths.TryGetValue(entry.Index, out string? qualified) ? qualified : entry.Name;
                        if (_streams.Count >= options.MaxStreamCount) throw new InvalidDataException("Compound stream count exceeds the configured limit.");
                        if (_streams.ContainsKey(path)) throw new InvalidDataException("Compound stream paths are ambiguous.");
                        if (entry.Size < 0 || entry.Size > options.MaxStreamBytes)
                            throw new InvalidDataException("Compound stream exceeds its configured byte limit.");
                        ValidateRegularStreamPhysicalBounds(path, entry.Size, _miniCutoff, physicalBytes);
                        total = checked(total + entry.Size);
                        if (total > options.MaxTotalStreamBytes) throw new InvalidDataException("Compound streams exceed the aggregate byte limit.");
                        options.StreamSizeValidator?.Invoke(path, entry.Size);
                        _streams.Add(path, entry);
                    }
                    _miniFat = miniFatStart == EndOfChain || miniFatCount == 0 ? Array.Empty<uint>()
                        : BytesToUInt32Array(ReadRegularChain(source, _basePosition, miniFatStart,
                            checked((long)miniFatCount * _sectorSize), _sectorSize, _physicalSectorCount, _fatSectorIds, _fatCache, token), token);
                    _rootChain = root.StartSector == EndOfChain || root.Size == 0 ? new List<uint>()
                        : GetRegularSectorChain(source, _basePosition, root.StartSector, root.Size, _sectorSize,
                            _physicalSectorCount, _fatSectorIds, _fatCache, token);
                } catch (Exception exception) when (exception is OverflowException || exception is IndexOutOfRangeException ||
                    exception is ArgumentException || exception is EndOfStreamException) {
                    throw new InvalidDataException("The compound index contains invalid bounds or truncated allocation metadata.", exception);
                } finally { source.Position = _basePosition; }
            }

            internal IReadOnlyList<OfficeCompoundFileEntry> Entries { get; }

            /// <summary>The physical root whose allocation and child tree this session uses.</summary>
            internal OfficeCompoundFileEntry RootEntry { get; }

            /// <summary>True when one active root exists at directory ID zero and is named Root Entry.</summary>
            internal bool HasCanonicalRootDirectory { get; }

            /// <summary>True for unreachable entries, cycles, shared nodes, invalid links or invalid entry roles.
            /// Compatibility lookups remain available; strict formats must reject this structure.</summary>
            internal bool HasInvalidDirectoryTree { get; }

            internal long GetStreamLength(string path) {
                ThrowIfDisposed(); _token.ThrowIfCancellationRequested();
                if (!_streams.TryGetValue(path, out DirectoryEntry? entry)) throw new InvalidDataException("Required compound stream is missing: " + path);
                return entry.Size;
            }

            /// <summary>Opens an exact, case-sensitive path without rescanning the directory or buffering its payload.</summary>
            internal Stream OpenStream(string path) {
                ThrowIfDisposed(); _token.ThrowIfCancellationRequested();
                if (!_streams.TryGetValue(path, out DirectoryEntry? entry)) throw new InvalidDataException("Required compound stream is missing: " + path);
                List<uint>? regular = null; List<uint>? mini = null;
                if (_validatedRegularChains != null && _validatedMiniChains != null) {
                    _validatedRegularChains.TryGetValue(path, out regular);
                    _validatedMiniChains.TryGetValue(path, out mini);
                } else if (entry.Size >= _miniCutoff) regular = GetRegularSectorChain(_source, _basePosition, entry.StartSector,
                    entry.Size, _sectorSize, _physicalSectorCount, _fatSectorIds, _fatCache, _token);
                else if (entry.Size > 0) mini = GetMiniSectorChain(entry.StartSector, entry.Size, _miniFat);
                return new CompoundEntryStream(_source, _basePosition, _sectorSize, entry.Size, regular, mini,
                    _rootChain, _rootSize, leaveOpen: true, _token, () => _disposed);
            }

            public void Dispose() {
                if (_disposed) return;
                _disposed = true; _streams.Clear(); _fatCache.Clear();
                if (!_leaveOpen) _source.Dispose();
            }
            private void ThrowIfDisposed() { if (_disposed) throw new ObjectDisposedException(nameof(ReadSession)); }
        }
    }
}
