using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Core.Internal {
    internal static partial class OfficeCompoundFileReader {
        internal sealed partial class ReadSession {
            /// <summary>Validates disjoint, exactly terminated allocation once for strict formats.
            /// This visits allocation metadata and every declared stream, without buffering payloads.
            /// Compatibility sessions retain their existing lookup and lazy-read behavior.</summary>
            internal void ValidateAllocation() {
                ThrowIfDisposed();
                _token.ThrowIfCancellationRequested();
                if (_validatedRegularChains != null) return;
                try {
                    byte[] header = ReadAt(_source, _basePosition, HeaderSize, _token);
                    var owned = new bool[_physicalSectorCount];
                    var miniOwned = new bool[checked((int)((_rootSize + MiniSectorSize - 1) / MiniSectorSize))];
                    var regular = new Dictionary<string, List<uint>>(StringComparer.Ordinal);
                    var mini = new Dictionary<string, List<uint>>(StringComparer.Ordinal);

                    uint Next(uint sector) => ReadFatEntry(_source, _basePosition, sector, _sectorSize,
                        _physicalSectorCount, _fatSectorIds, _fatCache, _token);
                    void Claim(uint sector) {
                        _token.ThrowIfCancellationRequested();
                        if (sector >= owned.Length || owned[sector])
                            throw new InvalidDataException("Compound allocation is shared, cyclic or outside the file.");
                        owned[sector] = true;
                    }
                    List<uint> Walk(uint start, long bytes, bool directory = false) {
                        long required = directory ? -1 : (bytes + _sectorSize - 1) / _sectorSize;
                        var chain = new List<uint>();
                        uint sector = start;
                        while (directory ? sector != EndOfChain : chain.Count < required) {
                            Claim(sector);
                            chain.Add(sector);
                            sector = Next(sector);
                        }
                        if (sector != EndOfChain)
                            throw new InvalidDataException("Compound allocation chain does not terminate at its declared length.");
                        return chain;
                    }

                    foreach (uint sector in _fatSectorIds) {
                        Claim(sector);
                        if (Next(sector) != 0xfffffffdU) throw new InvalidDataException("Compound FAT sector has an invalid allocation marker.");
                    }
                    uint difat = ReadUInt32(header, 68);
                    int difatCount = checked((int)ReadUInt32(header, 72));
                    for (int index = 0; index < difatCount; index++) {
                        Claim(difat);
                        if (Next(difat) != 0xfffffffcU) throw new InvalidDataException("Compound DIFAT sector has an invalid allocation marker.");
                        byte[] data = ReadSector(_source, _basePosition, difat, _sectorSize, _physicalSectorCount, _token);
                        difat = ReadUInt32(data, _sectorSize - 4);
                    }
                    if (difat != EndOfChain) throw new InvalidDataException("Compound DIFAT chain does not terminate at its declared count.");
                    List<uint> directory = Walk(ReadUInt32(header, 48), 0, directory: true);
                    uint declaredDirectorySectors = ReadUInt32(header, 40);
                    if (ReadUInt16(header, 26) == 3 ? declaredDirectorySectors != 0 : declaredDirectorySectors != directory.Count)
                        throw new InvalidDataException("Compound directory sector count is inconsistent with its format version or chain.");
                    Walk(ReadUInt32(header, 60), checked((long)ReadUInt32(header, 64) * _sectorSize));
                    // RootEntry identifies the same physical root used by the session's mini stream.
                    List<uint> root = Walk(_rootStartSector, _rootSize);
                    if (root.Count != _rootChain.Count) throw new InvalidDataException("Compound root mini-stream allocation changed.");

                    foreach (var pair in _streams) {
                        _token.ThrowIfCancellationRequested();
                        DirectoryEntry entry = pair.Value;
                        if (entry.Size >= _miniCutoff) {
                            regular.Add(pair.Key, Walk(entry.StartSector, entry.Size));
                        } else {
                            var chain = new List<uint>();
                            uint sector = entry.StartSector;
                            long remaining = entry.Size;
                            while (remaining > 0) {
                                _token.ThrowIfCancellationRequested();
                                if (sector >= miniOwned.Length || sector >= _miniFat.Length || miniOwned[sector]
                                    || (long)sector * MiniSectorSize > _rootSize - Math.Min(MiniSectorSize, remaining))
                                    throw new InvalidDataException("Compound mini allocation is shared, cyclic or outside its root stream.");
                                miniOwned[sector] = true;
                                chain.Add(sector);
                                sector = _miniFat[sector];
                                remaining -= Math.Min(MiniSectorSize, remaining);
                            }
                            if (sector != EndOfChain)
                                throw new InvalidDataException("Compound mini allocation chain does not terminate at its declared length.");
                            mini.Add(pair.Key, chain);
                        }
                    }
                    for (uint sector = 0; sector < owned.Length; sector++) {
                        _token.ThrowIfCancellationRequested();
                        if (!owned[sector] && Next(sector) != FreeSect)
                            throw new InvalidDataException("Compound file contains allocated sectors without an owner.");
                    }
                    for (int sector = 0; sector < _miniFat.Length; sector++) {
                        _token.ThrowIfCancellationRequested();
                        if ((sector >= miniOwned.Length || !miniOwned[sector]) && _miniFat[sector] != FreeSect)
                            throw new InvalidDataException("Compound mini stream contains allocated sectors without an owner.");
                    }
                    _validatedRegularChains = regular;
                    _validatedMiniChains = mini;
                } catch (Exception exception) when (exception is OverflowException || exception is ArgumentException
                    || exception is IndexOutOfRangeException || exception is EndOfStreamException) {
                    throw new InvalidDataException("Compound allocation metadata has invalid bounds.", exception);
                } finally { _source.Position = _basePosition; }
            }
        }
    }
}
