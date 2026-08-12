using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Provenance;

internal static class OfficeProvenanceTiff {
    private const ushort C2paTag = 0xCD41;
    private const ushort XmpTag = 700;
    private const ushort ByteType = 1;
    private const ushort UndefinedType = 7;

    internal static void Inspect(byte[] data, OfficeProvenanceOptions options, OfficeProvenanceContext context) {
        List<TiffIfd> ifds = ReadIfds(data, options);
        for (int ifdIndex = 0; ifdIndex < ifds.Count; ifdIndex++) {
            TiffIfd ifd = ifds[ifdIndex];
            foreach (TiffEntry entry in ifd.Entries) {
                if (entry.Tag == XmpTag && (entry.Type == ByteType || entry.Type == UndefinedType) &&
                    TryGetPayload(data, entry, options.MaxAssetBytes, out int xmpOffset, out int xmpLength)) {
                    byte[] packet = new byte[xmpLength];
                    Buffer.BlockCopy(data, xmpOffset, packet, 0, xmpLength);
                    OfficeProvenanceXmp.Inspect(packet, options, context, $"TIFF/IFD[{ifdIndex}]/XMP@{entry.Offset}");
                    continue;
                }
                if (entry.Tag != C2paTag) continue;
                bool valid = ifdIndex == ifds.Count - 1 && entry.Type == UndefinedType && TryGetPayload(data, entry, options.MaxManifestBytes, out int payloadOffset, out int payloadLength) &&
                    OfficeC2paManifestStore.IsValid(
                        data, payloadOffset, payloadLength, options.MaxManifestBytes, options.MaxContainerEntries, out _);
                string location = $"TIFF/IFD[{ifdIndex}]/0xCD41@{entry.Offset}";
                context.Add(new OfficeProvenanceEvidence(OfficeProvenanceCarrierKind.C2paManifest, location, valid, entry.Count > long.MaxValue ? long.MaxValue : (long)entry.Count));
                if (ifdIndex != ifds.Count - 1) context.Diagnostics.Add($"The C2PA TIFF tag at IFD {ifdIndex} is not in the last main IFD.");
            }
        }
    }

    internal static byte[] Remove(byte[] data, OfficeProvenanceRemovalOptions options, List<OfficeProvenanceChange> changes, out bool reserialized) {
        reserialized = false;
        if (!options.RemoveC2paManifests && !options.RemoveAiSourceMetadata) return (byte[])data.Clone();
        List<TiffIfd> ifds = ReadIfds(data, options.Limits);
        byte[] output = (byte[])data.Clone();
        for (int ifdIndex = 0; ifdIndex < ifds.Count; ifdIndex++) {
            TiffIfd ifd = ifds[ifdIndex];
            var retained = new List<TiffEntry>(ifd.Entries.Count);
            foreach (TiffEntry entry in ifd.Entries) {
                if (entry.Tag == XmpTag && (entry.Type == ByteType || entry.Type == UndefinedType) &&
                    TryGetPayload(data, entry, options.Limits.MaxAssetBytes, out int xmpOffset, out int xmpLength)) {
                    byte[] packet = new byte[xmpLength];
                    Buffer.BlockCopy(data, xmpOffset, packet, 0, xmpLength);
                    var pendingChanges = new List<OfficeProvenanceChange>();
                    if (OfficeProvenanceXmp.TryRemoveAiDeclarations(
                        packet,
                        options,
                        $"TIFF/IFD[{ifdIndex}]/XMP@{entry.Offset}",
                        pendingChanges,
                        out byte[] cleaned)) {
                        if (HasOverlappingValueStorage(data, ifds, entry, xmpOffset, xmpLength)) {
                            retained.Add(entry);
                            continue;
                        }
                        if (cleaned.Length > xmpLength) throw new InvalidDataException("Rewritten TIFF XMP exceeds its existing allocation.");
                        Buffer.BlockCopy(cleaned, 0, output, xmpOffset, cleaned.Length);
                        Array.Clear(output, xmpOffset + cleaned.Length, xmpLength - cleaned.Length);
                        WriteEntryCount(output, entry, (ulong)cleaned.Length);
                        changes.AddRange(pendingChanges);
                        reserialized = true;
                    }
                    retained.Add(entry);
                    continue;
                }
                if (entry.Tag != C2paTag) { retained.Add(entry); continue; }
                bool valid = ifdIndex == ifds.Count - 1 && entry.Type == UndefinedType && TryGetPayload(data, entry, options.Limits.MaxManifestBytes, out int payloadOffset, out int payloadLength) &&
                    OfficeC2paManifestStore.IsValid(
                        data, payloadOffset, payloadLength, options.Limits.MaxManifestBytes, options.Limits.MaxContainerEntries, out _);
                if (options.RemoveC2paManifests && (valid || !options.RequireStructurallyValidCarrier)) {
                    changes.Add(new OfficeProvenanceChange(
                        OfficeProvenanceCarrierKind.C2paManifest,
                        $"TIFF/IFD[{ifdIndex}]/0xCD41@{entry.Offset}",
                        removedBytes: 0));
                } else {
                    retained.Add(entry);
                }
            }
            if (retained.Count == ifd.Entries.Count) continue;
            RewriteIfd(output, ifd, retained);
            reserialized = true;
        }
        return output;
    }

    private static bool HasOverlappingValueStorage(
        byte[] data,
        List<TiffIfd> ifds,
        TiffEntry xmpEntry,
        int xmpOffset,
        int xmpLength) {
        long xmpEnd = (long)xmpOffset + xmpLength;
        foreach (TiffIfd ifd in ifds) {
            foreach (TiffEntry entry in ifd.Entries) {
                if (ReferenceEquals(entry, xmpEntry) || !TryGetValueStorageRange(data, entry, out int offset, out int length)) continue;
                long end = (long)offset + length;
                if (offset < xmpEnd && xmpOffset < end) return true;
            }
        }
        return false;
    }

    private static bool TryGetValueStorageRange(byte[] data, TiffEntry entry, out int offset, out int length) {
        offset = length = 0;
        int elementSize = entry.Type switch {
            1 or 2 or 6 or 7 => 1,
            3 or 8 => 2,
            4 or 9 or 11 or 13 => 4,
            5 or 10 or 12 or 16 or 17 or 18 => 8,
            _ => 0
        };
        if (elementSize == 0 || entry.Count == 0 || entry.Count > (ulong)(int.MaxValue / elementSize)) return false;
        length = (int)entry.Count * elementSize;
        if (length <= entry.InlineSize) {
            offset = entry.Offset + (entry.InlineSize == 8 ? 12 : 8);
            return offset <= data.Length - length;
        }
        if (entry.ValueOrOffset > int.MaxValue || entry.ValueOrOffset > (ulong)(data.Length - length)) return false;
        offset = (int)entry.ValueOrOffset;
        return true;
    }

    private static void WriteEntryCount(byte[] output, TiffEntry entry, ulong count) {
        if (entry.InlineSize == 8) OfficeProvenanceBinary.WriteUInt64(output, entry.Offset + 4, count, entry.LittleEndian);
        else OfficeProvenanceBinary.WriteUInt32(output, entry.Offset + 4, checked((uint)count), entry.LittleEndian);
    }

    private static List<TiffIfd> ReadIfds(byte[] data, OfficeProvenanceOptions options) {
        if (data.Length < 8) throw new InvalidDataException("TIFF header is truncated.");
        bool littleEndian;
        if (data[0] == (byte)'I' && data[1] == (byte)'I') littleEndian = true;
        else if (data[0] == (byte)'M' && data[1] == (byte)'M') littleEndian = false;
        else throw new InvalidDataException("TIFF byte-order marker is invalid.");

        ushort version = OfficeProvenanceBinary.ReadUInt16(data, 2, littleEndian);
        bool bigTiff = version == 43;
        ulong nextOffset;
        if (version == 42) {
            nextOffset = OfficeProvenanceBinary.ReadUInt32(data, 4, littleEndian);
        } else if (bigTiff) {
            if (data.Length < 16 || OfficeProvenanceBinary.ReadUInt16(data, 4, littleEndian) != 8 ||
                OfficeProvenanceBinary.ReadUInt16(data, 6, littleEndian) != 0) throw new InvalidDataException("BigTIFF header is invalid.");
            nextOffset = OfficeProvenanceBinary.ReadUInt64(data, 8, littleEndian);
        } else {
            throw new InvalidDataException("TIFF version is not supported.");
        }

        int countFieldSize = bigTiff ? 8 : 2;
        int entrySize = bigTiff ? 20 : 12;
        int nextFieldSize = bigTiff ? 8 : 4;
        var visited = new HashSet<ulong>();
        var result = new List<TiffIfd>();
        int totalEntryCount = 0;
        while (nextOffset != 0) {
            if (!visited.Add(nextOffset)) throw new InvalidDataException("TIFF main IFD chain contains a cycle.");
            if (visited.Count > 1024) throw new InvalidDataException("TIFF main IFD chain exceeds the supported count.");
            if (nextOffset > int.MaxValue || nextOffset > (ulong)(data.Length - countFieldSize)) throw new InvalidDataException("TIFF IFD offset exceeds the asset bounds.");
            int ifdOffset = (int)nextOffset;
            ulong countValue = bigTiff
                ? OfficeProvenanceBinary.ReadUInt64(data, ifdOffset, littleEndian)
                : OfficeProvenanceBinary.ReadUInt16(data, ifdOffset, littleEndian);
            if (countValue > 65535) throw new InvalidDataException("TIFF IFD entry count exceeds the supported limit.");
            int count = (int)countValue;
            if (count > options.MaxContainerEntries - totalEntryCount) {
                throw new InvalidDataException("TIFF IFD entries exceed the configured container-entry limit.");
            }
            totalEntryCount += count;
            long tableEndValue = (long)ifdOffset + countFieldSize + (long)count * entrySize + nextFieldSize;
            if (tableEndValue > data.Length) throw new InvalidDataException("TIFF IFD table exceeds the asset bounds.");
            int entriesOffset = ifdOffset + countFieldSize;
            var entries = new List<TiffEntry>(count);
            for (int index = 0; index < count; index++) {
                int entryOffset = entriesOffset + index * entrySize;
                ushort tag = OfficeProvenanceBinary.ReadUInt16(data, entryOffset, littleEndian);
                ushort type = OfficeProvenanceBinary.ReadUInt16(data, entryOffset + 2, littleEndian);
                ulong valueCount = bigTiff
                    ? OfficeProvenanceBinary.ReadUInt64(data, entryOffset + 4, littleEndian)
                    : OfficeProvenanceBinary.ReadUInt32(data, entryOffset + 4, littleEndian);
                ulong valueOrOffset = bigTiff
                    ? OfficeProvenanceBinary.ReadUInt64(data, entryOffset + 12, littleEndian)
                    : OfficeProvenanceBinary.ReadUInt32(data, entryOffset + 8, littleEndian);
                entries.Add(new TiffEntry(entryOffset, tag, type, valueCount, valueOrOffset, bigTiff ? 8 : 4, littleEndian));
            }
            int nextFieldOffset = entriesOffset + count * entrySize;
            nextOffset = bigTiff
                ? OfficeProvenanceBinary.ReadUInt64(data, nextFieldOffset, littleEndian)
                : OfficeProvenanceBinary.ReadUInt32(data, nextFieldOffset, littleEndian);
            result.Add(new TiffIfd(ifdOffset, entriesOffset, nextFieldOffset, countFieldSize, entrySize, nextFieldSize, littleEndian, bigTiff, entries));
        }
        return result;
    }

    private static bool TryGetPayload(byte[] data, TiffEntry entry, long maximumBytes, out int offset, out int length) {
        offset = length = 0;
        if (entry.Count == 0 || entry.Count > (ulong)maximumBytes || entry.Count > int.MaxValue) return false;
        length = (int)entry.Count;
        if (length <= entry.InlineSize) {
            offset = entry.Offset + (entry.InlineSize == 8 ? 12 : 8);
            return offset <= data.Length - length;
        }
        if (entry.ValueOrOffset > int.MaxValue || entry.ValueOrOffset > (ulong)(data.Length - length)) return false;
        offset = (int)entry.ValueOrOffset;
        return true;
    }

    private static void RewriteIfd(byte[] output, TiffIfd ifd, List<TiffEntry> retained) {
        if (ifd.BigTiff) OfficeProvenanceBinary.WriteUInt64(output, ifd.Offset, (ulong)retained.Count, ifd.LittleEndian);
        else OfficeProvenanceBinary.WriteUInt16(output, ifd.Offset, (ushort)retained.Count, ifd.LittleEndian);
        int target = ifd.EntriesOffset;
        foreach (TiffEntry entry in retained) {
            Buffer.BlockCopy(output, entry.Offset, output, target, ifd.EntrySize);
            target += ifd.EntrySize;
        }
        Buffer.BlockCopy(output, ifd.NextFieldOffset, output, target, ifd.NextFieldSize);
        target += ifd.NextFieldSize;
        int oldEnd = ifd.NextFieldOffset + ifd.NextFieldSize;
        Array.Clear(output, target, oldEnd - target);
    }

    private sealed class TiffEntry {
        internal TiffEntry(int offset, ushort tag, ushort type, ulong count, ulong valueOrOffset, int inlineSize, bool littleEndian) {
            Offset = offset; Tag = tag; Type = type; Count = count; ValueOrOffset = valueOrOffset; InlineSize = inlineSize; LittleEndian = littleEndian;
        }
        internal int Offset { get; }
        internal ushort Tag { get; }
        internal ushort Type { get; }
        internal ulong Count { get; }
        internal ulong ValueOrOffset { get; }
        internal int InlineSize { get; }
        internal bool LittleEndian { get; }
    }

    private sealed class TiffIfd {
        internal TiffIfd(int offset, int entriesOffset, int nextFieldOffset, int countFieldSize, int entrySize, int nextFieldSize,
            bool littleEndian, bool bigTiff, List<TiffEntry> entries) {
            Offset = offset; EntriesOffset = entriesOffset; NextFieldOffset = nextFieldOffset; CountFieldSize = countFieldSize;
            EntrySize = entrySize; NextFieldSize = nextFieldSize; LittleEndian = littleEndian; BigTiff = bigTiff; Entries = entries;
        }
        internal int Offset { get; }
        internal int EntriesOffset { get; }
        internal int NextFieldOffset { get; }
        internal int CountFieldSize { get; }
        internal int EntrySize { get; }
        internal int NextFieldSize { get; }
        internal bool LittleEndian { get; }
        internal bool BigTiff { get; }
        internal List<TiffEntry> Entries { get; }
    }
}
