using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Provenance;

internal static class OfficeProvenanceGif {
    internal static void Inspect(byte[] data, OfficeProvenanceOptions options, OfficeProvenanceContext context) {
        Walk(data, options, context, output: null, removalOptions: null, changes: null);
    }

    internal static byte[] Remove(byte[] data, OfficeProvenanceRemovalOptions options, List<OfficeProvenanceChange> changes) {
        if (!options.RemoveC2paManifests && !options.RemoveAiSourceMetadata) return (byte[])data.Clone();
        using var output = new MemoryStream(data.Length);
        int bodyOffset = GetBodyOffset(data);
        output.Write(data, 0, bodyOffset);
        Walk(data, options.Limits, context: null, output, options, changes);
        return output.ToArray();
    }

    private static void Walk(
        byte[] data,
        OfficeProvenanceOptions options,
        OfficeProvenanceContext? context,
        Stream? output,
        OfficeProvenanceRemovalOptions? removalOptions,
        List<OfficeProvenanceChange>? changes) {
        int offset = GetBodyOffset(data);
        bool foundTrailer = false;
        int entryCount = 0;
        while (offset < data.Length) {
            ReserveEntry(ref entryCount, options.MaxContainerEntries);
            int blockStart = offset;
            byte introducer = data[offset++];
            if (introducer == 0x3B) {
                output?.Write(data, blockStart, 1);
                foundTrailer = true;
                break;
            }
            if (introducer == 0x2C) {
                if (data.Length - offset < 9) throw new InvalidDataException("GIF image descriptor is truncated.");
                byte packed = data[offset + 8];
                offset += 9;
                if ((packed & 0x80) != 0) {
                    int tableBytes = 3 << ((packed & 0x07) + 1);
                    if (tableBytes > data.Length - offset) throw new InvalidDataException("GIF local color table is truncated.");
                    offset += tableBytes;
                }
                if (offset >= data.Length) throw new InvalidDataException("GIF image data is truncated.");
                offset++; // LZW minimum code size.
                offset = SkipSubBlocks(data, offset, options.MaxAssetBytes, ref entryCount, options.MaxContainerEntries, out _);
                output?.Write(data, blockStart, offset - blockStart);
                continue;
            }
            if (introducer != 0x21 || offset >= data.Length) throw new InvalidDataException("GIF contains an unsupported or truncated block.");
            byte label = data[offset++];
            if (label == 0xFF) {
                if (offset >= data.Length) throw new InvalidDataException("GIF application extension is truncated.");
                int headerLength = data[offset++];
                if (headerLength > data.Length - offset) throw new InvalidDataException("GIF application extension header is truncated.");
                bool isC2pa = headerLength == 11 && OfficeProvenanceBinary.MatchesAscii(data, offset, "C2PA_GIF") &&
                    data[offset + 8] == 0x01 && data[offset + 9] == 0x00 && data[offset + 10] == 0x00;
                bool isXmp = headerLength == 11 && OfficeProvenanceBinary.MatchesAscii(data, 0, "GIF89a") &&
                    OfficeProvenanceBinary.MatchesAscii(data, offset, "XMP DataXMP");
                offset += headerLength;
                int payloadStart = offset;
                if (isXmp && TryReadXmpApplicationData(
                    data, payloadStart, options.MaxAssetBytes, out int packetLength, out int extensionEnd)) {
                    ReserveEntry(ref entryCount, options.MaxContainerEntries);
                    byte[] packet = new byte[packetLength];
                    Buffer.BlockCopy(data, payloadStart, packet, 0, packetLength);
                    string location = $"GIF/XMP@{blockStart}";
                    if (context != null) OfficeProvenanceXmp.Inspect(packet, options, context, location);
                    if (output != null && removalOptions != null && changes != null && removalOptions.RemoveAiSourceMetadata &&
                        OfficeProvenanceXmp.TryRemoveAiDeclarations(packet, removalOptions, location, changes, out byte[] cleaned)) {
                        output.Write(data, blockStart, payloadStart - blockStart);
                        output.Write(cleaned, 0, cleaned.Length);
                        output.Write(data, payloadStart + packetLength, extensionEnd - payloadStart - packetLength);
                    } else {
                        output?.Write(data, blockStart, extensionEnd - blockStart);
                    }
                    offset = extensionEnd;
                    continue;
                }
                offset = SkipSubBlocks(data, offset, isC2pa ? options.MaxManifestBytes : options.MaxAssetBytes,
                    ref entryCount, options.MaxContainerEntries, out int payloadLength);
                if (isC2pa) {
                    byte[] manifest = CollectSubBlocks(data, payloadStart, payloadLength);
                    bool valid = OfficeC2paManifestStore.IsValid(
                        manifest, 0, manifest.Length, options.MaxManifestBytes, options.MaxContainerEntries, out _);
                    string location = $"GIF/C2PA_GIF@{blockStart}";
                    context?.Add(new OfficeProvenanceEvidence(OfficeProvenanceCarrierKind.C2paManifest, location, valid, manifest.Length));
                    bool remove = output != null && removalOptions != null && changes != null &&
                        removalOptions.RemoveC2paManifests &&
                        (valid || !removalOptions.RequireStructurallyValidCarrier);
                    if (remove) changes!.Add(new OfficeProvenanceChange(OfficeProvenanceCarrierKind.C2paManifest, location, offset - blockStart));
                    else output?.Write(data, blockStart, offset - blockStart);
                } else {
                    output?.Write(data, blockStart, offset - blockStart);
                }
            } else {
                offset = SkipSubBlocks(data, offset, options.MaxAssetBytes, ref entryCount, options.MaxContainerEntries, out _);
                output?.Write(data, blockStart, offset - blockStart);
            }
        }
        if (!foundTrailer) throw new InvalidDataException("GIF does not contain a trailer.");
        if (offset < data.Length) output?.Write(data, offset, data.Length - offset);
    }

    private static int GetBodyOffset(byte[] data) {
        if (data.Length < 13 || !(OfficeProvenanceBinary.MatchesAscii(data, 0, "GIF87a") || OfficeProvenanceBinary.MatchesAscii(data, 0, "GIF89a"))) {
            throw new InvalidDataException("GIF header is invalid.");
        }
        int offset = 13;
        byte packed = data[10];
        if ((packed & 0x80) != 0) {
            int tableBytes = 3 << ((packed & 0x07) + 1);
            if (tableBytes > data.Length - offset) throw new InvalidDataException("GIF global color table is truncated.");
            offset += tableBytes;
        }
        return offset;
    }

    private static int SkipSubBlocks(
        byte[] data,
        int offset,
        long maximumPayload,
        ref int entryCount,
        int maximumEntries,
        out int payloadLength) {
        long total = 0;
        while (true) {
            ReserveEntry(ref entryCount, maximumEntries);
            if (offset >= data.Length) throw new InvalidDataException("GIF data sub-blocks are truncated.");
            int length = data[offset++];
            if (length == 0) break;
            if (length > data.Length - offset) throw new InvalidDataException("GIF data sub-block exceeds the asset bounds.");
            total += length;
            if (total > maximumPayload || total > int.MaxValue) throw new InvalidDataException("GIF data sub-blocks exceed the configured limit.");
            offset += length;
        }
        payloadLength = (int)total;
        return offset;
    }

    private static void ReserveEntry(ref int entryCount, int maximumEntries) {
        if (entryCount >= maximumEntries) {
            throw new InvalidDataException($"The GIF exceeds the configured container entry limit of {maximumEntries}.");
        }
        entryCount++;
    }

    private static byte[] CollectSubBlocks(byte[] data, int offset, int payloadLength) {
        byte[] result = new byte[payloadLength];
        int target = 0;
        while (target < result.Length) {
            int length = data[offset++];
            Buffer.BlockCopy(data, offset, result, target, length);
            offset += length;
            target += length;
        }
        return result;
    }

    private static bool TryReadXmpApplicationData(
        byte[] data,
        int payloadOffset,
        long maximumPacketBytes,
        out int packetLength,
        out int extensionEnd) {
        packetLength = 0;
        extensionEnd = payloadOffset;
        const int trailerLength = 258;
        for (int candidate = payloadOffset; candidate <= data.Length - trailerLength; candidate++) {
            if (data[candidate] != 0x01 || data[candidate + 1] != 0xFF) continue;
            bool valid = true;
            for (int index = 1; index <= 255; index++) {
                if (data[candidate + index] == 256 - index) continue;
                valid = false;
                break;
            }
            if (!valid || data[candidate + 256] != 0 || data[candidate + 257] != 0) continue;
            long length = candidate - (long)payloadOffset;
            if (length <= 0 || length > maximumPacketBytes || length > int.MaxValue) return false;
            packetLength = (int)length;
            extensionEnd = candidate + trailerLength;
            return true;
        }
        return false;
    }
}
