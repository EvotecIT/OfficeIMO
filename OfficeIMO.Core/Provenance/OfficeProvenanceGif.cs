using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Provenance;

internal static class OfficeProvenanceGif {
    internal static void Inspect(byte[] data, OfficeProvenanceOptions options, OfficeProvenanceContext context) {
        Walk(data, options, context, output: null, removalOptions: null, changes: null);
    }

    internal static byte[] Remove(byte[] data, OfficeProvenanceRemovalOptions options, List<OfficeProvenanceChange> changes) {
        if (!options.RemoveC2paManifests) return (byte[])data.Clone();
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
        while (offset < data.Length) {
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
                offset = SkipSubBlocks(data, offset, options.MaxAssetBytes, out _);
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
                offset += headerLength;
                int payloadStart = offset;
                offset = SkipSubBlocks(data, offset, options.MaxManifestBytes, out int payloadLength);
                if (isC2pa) {
                    byte[] manifest = CollectSubBlocks(data, payloadStart, payloadLength);
                    bool valid = OfficeC2paManifestStore.IsValid(manifest, 0, manifest.Length, options.MaxManifestBytes, out _);
                    string location = $"GIF/C2PA_GIF@{blockStart}";
                    context?.Add(new OfficeProvenanceEvidence(OfficeProvenanceCarrierKind.C2paManifest, location, valid, manifest.Length));
                    bool remove = output != null && removalOptions != null && changes != null &&
                        (valid || !removalOptions.RequireStructurallyValidCarrier);
                    if (remove) changes!.Add(new OfficeProvenanceChange(OfficeProvenanceCarrierKind.C2paManifest, location, offset - blockStart));
                    else output?.Write(data, blockStart, offset - blockStart);
                } else {
                    output?.Write(data, blockStart, offset - blockStart);
                }
            } else {
                offset = SkipSubBlocks(data, offset, options.MaxAssetBytes, out _);
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

    private static int SkipSubBlocks(byte[] data, int offset, long maximumPayload, out int payloadLength) {
        long total = 0;
        while (true) {
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
}
