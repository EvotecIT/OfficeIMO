using System;
using System.Collections.Generic;
using System.IO;
using OfficeIMO.Drawing;

namespace OfficeIMO.Provenance;

internal static class OfficeProvenanceRiff {
    internal static void Inspect(byte[] data, OfficeProvenanceOptions options, OfficeProvenanceContext context) {
        Walk(data, options, context, output: null, removalOptions: null, changes: null);
    }

    internal static byte[] Remove(byte[] data, OfficeProvenanceRemovalOptions options, List<OfficeProvenanceChange> changes, out bool reserialized) {
        reserialized = false;
        if (!options.RemoveC2paManifests && !options.RemoveAiSourceMetadata) {
            return OfficeProvenanceBinary.CloneForOutput(data, options.EffectiveMaxOutputBytes);
        }
        using var output = new OfficeProvenanceBoundedMemoryStream(options.EffectiveMaxOutputBytes, data.Length);
        output.Write(data, 0, 12);
        reserialized = Walk(data, options.Limits, context: null, output, options, changes);
        int sourceDeclaredEnd = checked((int)(8L + OfficeProvenanceBinary.ReadUInt32(data, 4, littleEndian: true)));
        int suffixLength = data.Length - sourceDeclaredEnd;
        byte[] result = output.ToArray();
        int rewrittenDeclaredEnd = result.Length - suffixLength;
        OfficeProvenanceBinary.WriteUInt32(result, 4, (uint)(rewrittenDeclaredEnd - 8), littleEndian: true);
        return result;
    }

    private static bool Walk(
        byte[] data,
        OfficeProvenanceOptions options,
        OfficeProvenanceContext? context,
        Stream? output,
        OfficeProvenanceRemovalOptions? removalOptions,
        List<OfficeProvenanceChange>? changes) {
        bool reserialized = false;
        if (data.Length < 12 || !OfficeProvenanceBinary.MatchesAscii(data, 0, "RIFF") ||
            !OfficeProvenanceBinary.MatchesAscii(data, 8, "WEBP")) throw new InvalidDataException("WebP RIFF header is invalid.");
        uint riffSize = OfficeProvenanceBinary.ReadUInt32(data, 4, littleEndian: true);
        long declaredEndValue = 8L + riffSize;
        if (declaredEndValue < 12 || declaredEndValue != data.Length) throw new InvalidDataException("RIFF size does not match the asset bounds.");
        int declaredEnd = (int)declaredEndValue;
        int c2paChunkCount = CountChunks(data, declaredEnd, options.MaxContainerEntries, "C2PA");
        int xmpChunkCount = CountChunks(data, declaredEnd, options.MaxContainerEntries, "XMP ");
        int extendedHeaderCount = CountChunks(data, declaredEnd, options.MaxContainerEntries, "VP8X");
        int lastImagePayloadOffset = FindLastImagePayloadOffset(data, declaredEnd, options.MaxContainerEntries);
        bool allChunksHaveValidPadding = HaveValidChunkPadding(data, declaredEnd, options.MaxContainerEntries);
        bool extendedHeaderFeaturesAreConsistent = HaveConsistentExtendedHeaderFeatures(
            data, declaredEnd, options.MaxContainerEntries);
        bool hasValidWebpContainer = OfficeImageReader.TryValidateWebpContainer(data);
        int offset = 12;
        int chunkCount = 0;
        bool hasValidExtendedHeader = false;
        bool extendedHeaderAdvertisesXmp = false;
        int lossyImagePayloads = 0;
        int losslessImagePayloads = 0;
        int animationFramePayloads = 0;
        bool foundXmp = false;
        while (offset < declaredEnd) {
            if (++chunkCount > options.MaxContainerEntries) {
                throw OfficeProvenanceLimitException.Create("WebP exceeds the configured container entry limit.");
            }
            if (declaredEnd - offset < 8) throw new InvalidDataException("RIFF contains a truncated chunk header.");
            uint payloadValue = OfficeProvenanceBinary.ReadUInt32(data, offset + 4, littleEndian: true);
            if (payloadValue > int.MaxValue) throw new InvalidDataException("RIFF chunk exceeds the supported size.");
            int payloadLength = (int)payloadValue;
            long totalValue = 8L + payloadLength + (payloadLength & 1);
            if (totalValue > declaredEnd - offset) throw new InvalidDataException("RIFF chunk exceeds the declared container bounds.");
            int total = (int)totalValue;
            string chunkType = System.Text.Encoding.ASCII.GetString(data, offset, 4);
            if ((payloadLength & 1) != 0 && data[offset + 8 + payloadLength] != 0) {
                context?.Diagnostics.Add($"The WebP {chunkType} chunk has a nonzero padding byte.");
            }
            if (chunkCount == 1 && chunkType == "VP8X") {
                hasValidExtendedHeader = IsValidExtendedHeader(data, offset + 8, payloadLength);
                extendedHeaderAdvertisesXmp = hasValidExtendedHeader && (data[offset + 8] & 0x04) != 0;
                output?.Write(data, offset, total);
            } else if (chunkType == "VP8X") {
                hasValidExtendedHeader = false;
                extendedHeaderAdvertisesXmp = false;
                output?.Write(data, offset, total);
            } else if (chunkType == "C2PA") {
                if (payloadLength > options.MaxManifestBytes) throw OfficeProvenanceLimitException.Create("RIFF provenance chunk exceeds the configured manifest limit.");
                bool isLast = offset + total == declaredEnd;
                bool hasUnambiguousImagePayload = HasUnambiguousImagePayload(
                    lossyImagePayloads, losslessImagePayloads, animationFramePayloads);
                bool valid = c2paChunkCount == 1 &&
                    extendedHeaderCount == 1 && hasValidExtendedHeader &&
                    extendedHeaderFeaturesAreConsistent && allChunksHaveValidPadding && hasValidWebpContainer && isLast &&
                    hasUnambiguousImagePayload && OfficeC2paManifestStore.IsValid(
                    data, offset + 8, payloadLength, options.MaxManifestBytes, options.MaxContainerEntries, out _);
                string location = $"RIFF/C2PA@{offset}";
                context?.Add(new OfficeProvenanceEvidence(OfficeProvenanceCarrierKind.C2paManifest, location, valid, payloadLength));
                if (c2paChunkCount > 1) context?.Diagnostics.Add("The WebP container contains multiple C2PA chunks.");
                if (!isLast) context?.Diagnostics.Add("The C2PA chunk is not the last chunk in the first RIFF container.");
                bool remove = output != null && removalOptions != null && changes != null &&
                    removalOptions.RemoveC2paManifests &&
                    (valid || !removalOptions.RequireStructurallyValidCarrier);
                if (remove) {
                    changes!.Add(new OfficeProvenanceChange(OfficeProvenanceCarrierKind.C2paManifest, location, total));
                    reserialized = true;
                } else {
                    output?.Write(data, offset, total);
                }
            } else if (chunkType == "XMP ") {
                byte[] packet = new byte[payloadLength];
                Buffer.BlockCopy(data, offset + 8, packet, 0, payloadLength);
                string location = $"WebP/XMP@{offset}";
                bool carrierValid = xmpChunkCount == 1 && extendedHeaderCount == 1 &&
                    extendedHeaderFeaturesAreConsistent && allChunksHaveValidPadding && hasValidWebpContainer &&
                    hasValidExtendedHeader && extendedHeaderAdvertisesXmp &&
                    HasUnambiguousImagePayload(lossyImagePayloads, losslessImagePayloads, animationFramePayloads) &&
                    offset > lastImagePayloadOffset && !foundXmp;
                if (xmpChunkCount > 1) context?.Diagnostics.Add("The WebP container contains multiple XMP chunks.");
                if (context != null) OfficeProvenanceXmp.Inspect(packet, options, context, location, carrierValid);
                if (output != null && removalOptions != null && changes != null &&
                    (carrierValid || !removalOptions.RequireStructurallyValidCarrier) &&
                    OfficeProvenanceXmp.TryRemoveAiDeclarations(packet, removalOptions, location, changes, out byte[] cleaned)) {
                    WriteChunk(output, "XMP ", cleaned);
                    reserialized = true;
                } else {
                    output?.Write(data, offset, total);
                }
                foundXmp = true;
            } else {
                output?.Write(data, offset, total);
            }
            if (chunkType == "VP8 ") lossyImagePayloads++;
            else if (chunkType == "VP8L") losslessImagePayloads++;
            else if (chunkType == "ANMF") animationFramePayloads++;
            offset += total;
        }
        if (offset != declaredEnd) throw new InvalidDataException("RIFF chunks do not end on the declared boundary.");
        if (declaredEnd < data.Length) output?.Write(data, declaredEnd, data.Length - declaredEnd);
        return reserialized;
    }

    private static bool HasUnambiguousImagePayload(int lossy, int lossless, int animationFrames) =>
        lossy == 1 && lossless == 0 && animationFrames == 0 ||
        lossless == 1 && lossy == 0 && animationFrames == 0 ||
        animationFrames > 0 && lossy == 0 && lossless == 0;

    private static int CountChunks(byte[] data, int declaredEnd, int maximumEntries, string expectedType) {
        int matches = 0;
        int entries = 0;
        int offset = 12;
        while (offset < declaredEnd) {
            if (++entries > maximumEntries) throw OfficeProvenanceLimitException.Create("WebP exceeds the configured container entry limit.");
            if (declaredEnd - offset < 8) throw new InvalidDataException("RIFF contains a truncated chunk header.");
            uint payloadValue = OfficeProvenanceBinary.ReadUInt32(data, offset + 4, littleEndian: true);
            if (payloadValue > int.MaxValue) throw new InvalidDataException("RIFF chunk exceeds the supported size.");
            int payloadLength = (int)payloadValue;
            long totalValue = 8L + payloadLength + (payloadLength & 1);
            if (totalValue > declaredEnd - offset) throw new InvalidDataException("RIFF chunk exceeds the declared container bounds.");
            if (OfficeProvenanceBinary.MatchesAscii(data, offset, expectedType)) matches++;
            offset += (int)totalValue;
        }
        if (offset != declaredEnd) throw new InvalidDataException("RIFF chunks do not end on the declared boundary.");
        return matches;
    }

    private static int FindLastImagePayloadOffset(byte[] data, int declaredEnd, int maximumEntries) {
        int lastOffset = -1;
        int entries = 0;
        int offset = 12;
        while (offset < declaredEnd) {
            if (++entries > maximumEntries) throw OfficeProvenanceLimitException.Create("WebP exceeds the configured container entry limit.");
            if (declaredEnd - offset < 8) throw new InvalidDataException("RIFF contains a truncated chunk header.");
            uint payloadValue = OfficeProvenanceBinary.ReadUInt32(data, offset + 4, littleEndian: true);
            if (payloadValue > int.MaxValue) throw new InvalidDataException("RIFF chunk exceeds the supported size.");
            int payloadLength = (int)payloadValue;
            long totalValue = 8L + payloadLength + (payloadLength & 1);
            if (totalValue > declaredEnd - offset) throw new InvalidDataException("RIFF chunk exceeds the declared container bounds.");
            if (OfficeProvenanceBinary.MatchesAscii(data, offset, "VP8 ") ||
                OfficeProvenanceBinary.MatchesAscii(data, offset, "VP8L") ||
                OfficeProvenanceBinary.MatchesAscii(data, offset, "ANMF")) lastOffset = offset;
            offset += (int)totalValue;
        }
        if (offset != declaredEnd) throw new InvalidDataException("RIFF chunks do not end on the declared boundary.");
        return lastOffset;
    }

    private static bool HaveValidChunkPadding(byte[] data, int declaredEnd, int maximumEntries) {
        int entries = 0;
        int offset = 12;
        bool valid = true;
        while (offset < declaredEnd) {
            if (++entries > maximumEntries) throw OfficeProvenanceLimitException.Create("WebP exceeds the configured container entry limit.");
            if (declaredEnd - offset < 8) throw new InvalidDataException("RIFF contains a truncated chunk header.");
            uint payloadValue = OfficeProvenanceBinary.ReadUInt32(data, offset + 4, littleEndian: true);
            if (payloadValue > int.MaxValue) throw new InvalidDataException("RIFF chunk exceeds the supported size.");
            int payloadLength = (int)payloadValue;
            long totalValue = 8L + payloadLength + (payloadLength & 1);
            if (totalValue > declaredEnd - offset) throw new InvalidDataException("RIFF chunk exceeds the declared container bounds.");
            if ((payloadLength & 1) != 0 && data[offset + 8 + payloadLength] != 0) valid = false;
            offset += (int)totalValue;
        }
        if (offset != declaredEnd) throw new InvalidDataException("RIFF chunks do not end on the declared boundary.");
        return valid;
    }

    private static bool IsValidExtendedHeader(byte[] data, int payloadOffset, int payloadLength) {
        if (payloadLength != 10) return false;
        byte flags = data[payloadOffset];
        return (flags & 0xC1) == 0 &&
            data[payloadOffset + 1] == 0 && data[payloadOffset + 2] == 0 && data[payloadOffset + 3] == 0;
    }

    private static bool HaveConsistentExtendedHeaderFeatures(byte[] data, int declaredEnd, int maximumEntries) {
        if (declaredEnd < 30 || !OfficeProvenanceBinary.MatchesAscii(data, 12, "VP8X")) return false;
        uint headerLength = OfficeProvenanceBinary.ReadUInt32(data, 16, littleEndian: true);
        if (headerLength != 10 || !IsValidExtendedHeader(data, 20, 10)) return false;
        byte flags = data[20];
        int animationControls = CountChunks(data, declaredEnd, maximumEntries, "ANIM");
        int animationFrames = CountChunks(data, declaredEnd, maximumEntries, "ANMF");
        int xmpChunks = CountChunks(data, declaredEnd, maximumEntries, "XMP ");
        int exifChunks = CountChunks(data, declaredEnd, maximumEntries, "EXIF");
        int colorProfiles = CountChunks(data, declaredEnd, maximumEntries, "ICCP");
        bool declaresAnimation = (flags & 0x02) != 0;
        return declaresAnimation == (animationControls == 1 && animationFrames > 0) &&
            ((flags & 0x04) != 0) == (xmpChunks == 1) &&
            ((flags & 0x08) != 0) == (exifChunks == 1) &&
            ((flags & 0x20) != 0) == (colorProfiles == 1);
    }

    private static void WriteChunk(Stream output, string type, byte[] payload) {
        byte[] header = new byte[8];
        System.Text.Encoding.ASCII.GetBytes(type).CopyTo(header, 0);
        OfficeProvenanceBinary.WriteUInt32(header, 4, (uint)payload.Length, littleEndian: true);
        output.Write(header, 0, header.Length);
        output.Write(payload, 0, payload.Length);
        if ((payload.Length & 1) != 0) output.WriteByte(0);
    }
}
