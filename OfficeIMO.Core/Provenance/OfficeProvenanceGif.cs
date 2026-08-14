using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Provenance;

internal static class OfficeProvenanceGif {
    internal static void Inspect(byte[] data, OfficeProvenanceOptions options, OfficeProvenanceContext context) {
        Walk(data, options, context, output: null, removalOptions: null, changes: null, out _);
    }

    internal static byte[] Remove(
        byte[] data,
        OfficeProvenanceRemovalOptions options,
        List<OfficeProvenanceChange> changes,
        out bool reserialized) {
        reserialized = false;
        if (!options.RemoveC2paManifests && !options.RemoveAiSourceMetadata) return (byte[])data.Clone();
        using var output = new MemoryStream(data.Length);
        int bodyOffset = GetBodyOffset(data);
        output.Write(data, 0, bodyOffset);
        Walk(data, options.Limits, context: null, output, options, changes, out reserialized);
        return output.ToArray();
    }

    private static void Walk(
        byte[] data,
        OfficeProvenanceOptions options,
        OfficeProvenanceContext? context,
        Stream? output,
        OfficeProvenanceRemovalOptions? removalOptions,
        List<OfficeProvenanceChange>? changes,
        out bool reserialized) {
        reserialized = false;
        int offset = GetBodyOffset(data);
        int c2paApplicationCount = CountC2paApplications(
            data, offset, options, out int xmpApplicationCount, out bool validStructure);
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
                bool isGif89a = OfficeProvenanceBinary.MatchesAscii(data, 0, "GIF89a");
                bool isC2paApplication = headerLength == 11 && OfficeProvenanceBinary.MatchesAscii(data, offset, "C2PA_GIF") &&
                    data[offset + 8] == 0x01 && data[offset + 9] == 0x00 && data[offset + 10] == 0x00;
                bool isXmp = headerLength == 11 && isGif89a &&
                    OfficeProvenanceBinary.MatchesAscii(data, offset, "XMP DataXMP");
                offset += headerLength;
                int payloadStart = offset;
                if (isXmp && TryReadXmpApplicationData(
                    data, payloadStart, options.MaxAssetBytes, ref entryCount, options.MaxContainerEntries,
                    out byte[] packet, out int extensionEnd, out int trailerStart, out bool usesSubBlocks)) {
                    string location = $"GIF/XMP@{blockStart}";
                    bool carrierValid = xmpApplicationCount == 1 && validStructure;
                    if (context != null) OfficeProvenanceXmp.Inspect(packet, options, context, location, carrierValid);
                    if (output != null && removalOptions != null && changes != null && removalOptions.RemoveAiSourceMetadata &&
                        (carrierValid || !removalOptions.RequireStructurallyValidCarrier) &&
                        OfficeProvenanceXmp.TryRemoveAiDeclarations(packet, removalOptions, location, changes, out byte[] cleaned)) {
                        output.Write(data, blockStart, payloadStart - blockStart);
                        if (usesSubBlocks) WriteSubBlocks(output, cleaned);
                        else output.Write(cleaned, 0, cleaned.Length);
                        output.Write(data, trailerStart, extensionEnd - trailerStart);
                        reserialized = true;
                    } else {
                        output?.Write(data, blockStart, extensionEnd - blockStart);
                    }
                    offset = extensionEnd;
                    continue;
                }
                offset = SkipSubBlocks(data, offset, isC2paApplication ? options.MaxManifestBytes : options.MaxAssetBytes,
                    ref entryCount, options.MaxContainerEntries, out int payloadLength);
                if (isC2paApplication) {
                    byte[] manifest = CollectSubBlocks(data, payloadStart, payloadLength);
                    bool valid = c2paApplicationCount == 1 && validStructure && isGif89a && OfficeC2paManifestStore.IsValid(
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

    private static int CountC2paApplications(
        byte[] data,
        int offset,
        OfficeProvenanceOptions options,
        out int xmpApplicationCount,
        out bool validStructure) {
        int count = 0;
        xmpApplicationCount = 0;
        validStructure = false;
        int entryCount = 0;
        while (offset < data.Length) {
            ReserveEntry(ref entryCount, options.MaxContainerEntries);
            byte introducer = data[offset++];
            if (introducer == 0x3B) {
                validStructure = offset == data.Length;
                return count;
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
                offset++;
                offset = SkipSubBlocks(data, offset, options.MaxAssetBytes, ref entryCount, options.MaxContainerEntries, out _);
                continue;
            }
            if (introducer != 0x21 || offset >= data.Length) throw new InvalidDataException("GIF contains an unsupported or truncated block.");
            byte label = data[offset++];
            if (label == 0xFF) {
                if (offset >= data.Length) throw new InvalidDataException("GIF application extension is truncated.");
                int headerLength = data[offset++];
                if (headerLength > data.Length - offset) throw new InvalidDataException("GIF application extension header is truncated.");
                bool isC2paApplication = headerLength == 11 && OfficeProvenanceBinary.MatchesAscii(data, offset, "C2PA_GIF") &&
                    data[offset + 8] == 0x01 && data[offset + 9] == 0x00 && data[offset + 10] == 0x00;
                bool isXmp = headerLength == 11 && OfficeProvenanceBinary.MatchesAscii(data, 0, "GIF89a") &&
                    OfficeProvenanceBinary.MatchesAscii(data, offset, "XMP DataXMP");
                if (isC2paApplication) count++;
                if (isXmp) xmpApplicationCount++;
                offset += headerLength;
                if (isXmp && TryReadXmpApplicationData(
                    data, offset, options.MaxAssetBytes, ref entryCount, options.MaxContainerEntries,
                    out _, out int extensionEnd, out _, out _)) {
                    offset = extensionEnd;
                    continue;
                }
                offset = SkipSubBlocks(data, offset, isC2paApplication ? options.MaxManifestBytes : options.MaxAssetBytes,
                    ref entryCount, options.MaxContainerEntries, out _);
            } else {
                offset = SkipSubBlocks(data, offset, options.MaxAssetBytes, ref entryCount, options.MaxContainerEntries, out _);
            }
        }
        throw new InvalidDataException("GIF does not contain a trailer.");
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
        ref int entryCount,
        int maximumEntries,
        out byte[] packet,
        out int extensionEnd,
        out int trailerStart,
        out bool usesSubBlocks) {
        packet = Array.Empty<byte>();
        extensionEnd = payloadOffset;
        trailerStart = payloadOffset;
        usesSubBlocks = false;
        int cursor = payloadOffset;
        if (cursor >= data.Length || HasXmpMagicTrailer(data, cursor)) return false;
        if (LooksLikeRawXmpPacket(data, payloadOffset)) {
            return TryReadRawXmpApplicationData(
                data, payloadOffset, maximumPacketBytes, ref entryCount, maximumEntries,
                out packet, out extensionEnd, out trailerStart);
        }
        int candidateEntryCount = entryCount;
        using (var collected = new MemoryStream()) {
            while (cursor < data.Length) {
                if (HasXmpMagicTrailer(data, cursor)) {
                    if (collected.Length == 0) break;
                    packet = collected.ToArray();
                    extensionEnd = cursor + 258;
                    trailerStart = cursor;
                    usesSubBlocks = true;
                    entryCount = candidateEntryCount;
                    return true;
                }
                ReserveEntry(ref candidateEntryCount, maximumEntries);
                int length = data[cursor++];
                if (length == 0) return false;
                if (length > data.Length - cursor || collected.Length > maximumPacketBytes - length) {
                    throw new InvalidDataException("GIF XMP data sub-blocks are truncated or exceed the configured asset limit.");
                }
                collected.Write(data, cursor, length);
                cursor += length;
            }
        }
        return false;
    }

    private static bool LooksLikeRawXmpPacket(byte[] data, int offset) {
        if (offset <= data.Length - 3 && data[offset] == 0xEF && data[offset + 1] == 0xBB && data[offset + 2] == 0xBF) {
            offset += 3;
        }
        while (offset < data.Length && data[offset] is 0x09 or 0x0A or 0x0D or 0x20) offset++;
        if (offset >= data.Length || data[offset++] != (byte)'<' || offset >= data.Length) return false;
        byte next = data[offset];
        return next is (byte)'?' or (byte)'!' or (byte)'_' or (>= (byte)'A' and <= (byte)'Z') or (>= (byte)'a' and <= (byte)'z');
    }

    private static bool TryReadRawXmpApplicationData(
        byte[] data,
        int payloadOffset,
        long maximumPacketBytes,
        ref int entryCount,
        int maximumEntries,
        out byte[] packet,
        out int extensionEnd,
        out int trailerStart) {
        packet = Array.Empty<byte>();
        extensionEnd = payloadOffset;
        trailerStart = payloadOffset;
        const int trailerLength = 258;
        for (int candidate = payloadOffset; candidate <= data.Length - trailerLength; candidate++) {
            if (!HasXmpMagicTrailer(data, candidate)) continue;
            long length = candidate - (long)payloadOffset;
            if (length <= 0 || length > maximumPacketBytes || length > int.MaxValue) return false;
            int rawEntries = entryCount;
            ReserveEntry(ref rawEntries, maximumEntries);
            packet = new byte[(int)length];
            Buffer.BlockCopy(data, payloadOffset, packet, 0, packet.Length);
            extensionEnd = candidate + trailerLength;
            trailerStart = candidate;
            entryCount = rawEntries;
            return true;
        }
        return false;
    }

    private static bool HasXmpMagicTrailer(byte[] data, int offset) {
        const int trailerLength = 258;
        if (offset < 0 || offset > data.Length - trailerLength || data[offset] != 0x01 || data[offset + 1] != 0xFF) return false;
        for (int index = 1; index <= 255; index++) {
            if (data[offset + index] != 256 - index) return false;
        }
        return data[offset + 256] == 0 && data[offset + 257] == 0;
    }

    private static void WriteSubBlocks(Stream output, byte[] payload) {
        int offset = 0;
        while (offset < payload.Length) {
            int length = Math.Min(255, payload.Length - offset);
            output.WriteByte((byte)length);
            output.Write(payload, offset, length);
            offset += length;
        }
    }
}
