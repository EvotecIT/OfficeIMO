using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Provenance;

internal static class OfficeProvenancePng {
    private const int SignatureLength = 8;

    internal static void Inspect(byte[] data, OfficeProvenanceOptions options, OfficeProvenanceContext context) {
        Walk(data, options, context, output: null, removalOptions: null, changes: null);
    }

    internal static byte[] Remove(byte[] data, OfficeProvenanceRemovalOptions options, List<OfficeProvenanceChange> changes, out bool reserialized) {
        reserialized = false;
        if (!options.RemoveC2paManifests && !options.RemoveAiSourceMetadata) return (byte[])data.Clone();
        using var output = new MemoryStream(data.Length);
        output.Write(data, 0, SignatureLength);
        reserialized = Walk(data, options.Limits, context: null, output, options, changes);
        return output.ToArray();
    }

    private static bool Walk(
        byte[] data,
        OfficeProvenanceOptions options,
        OfficeProvenanceContext? context,
        Stream? output,
        OfficeProvenanceRemovalOptions? removalOptions,
        List<OfficeProvenanceChange>? changes) {
        bool reserialized = false;
        int offset = SignatureLength;
        bool foundEnd = false;
        while (offset < data.Length) {
            if (data.Length - offset < 12) throw new InvalidDataException("PNG contains a truncated chunk.");
            uint payloadValue = OfficeProvenanceBinary.ReadUInt32(data, offset, littleEndian: false);
            if (payloadValue > int.MaxValue || payloadValue > options.MaxManifestBytes && OfficeProvenanceBinary.MatchesAscii(data, offset + 4, "caBX")) {
                throw new InvalidDataException("PNG provenance chunk exceeds the configured manifest limit.");
            }
            int payloadLength = checked((int)payloadValue);
            long totalValue = 12L + payloadLength;
            if (totalValue > data.Length - offset) throw new InvalidDataException("PNG chunk length exceeds the remaining asset.");
            int total = (int)totalValue;
            bool isC2pa = OfficeProvenanceBinary.MatchesAscii(data, offset + 4, "caBX");
            if (isC2pa) {
                bool valid = HasValidCrc(data, offset, payloadLength) &&
                    OfficeC2paManifestStore.IsValid(data, offset + 8, payloadLength, options.MaxManifestBytes, out _);
                string location = $"PNG/caBX@{offset}";
                context?.Add(new OfficeProvenanceEvidence(OfficeProvenanceCarrierKind.C2paManifest, location, valid, payloadLength));
                bool remove = output != null && removalOptions != null && changes != null && removalOptions.RemoveC2paManifests &&
                    (valid || !removalOptions.RequireStructurallyValidCarrier);
                if (remove) changes!.Add(new OfficeProvenanceChange(OfficeProvenanceCarrierKind.C2paManifest, location, total));
                else output?.Write(data, offset, total);
            } else if (OfficeProvenanceBinary.MatchesAscii(data, offset + 4, "iTXt") &&
                TryGetXmpPacket(data, offset + 8, payloadLength, out int packetOffset, out int packetLength)) {
                byte[] packet = new byte[packetLength];
                Buffer.BlockCopy(data, packetOffset, packet, 0, packetLength);
                string location = $"PNG/iTXt-XMP@{offset}";
                if (context != null) OfficeProvenanceXmp.Inspect(packet, options, context, location);
                if (output != null && removalOptions != null && changes != null &&
                    OfficeProvenanceXmp.TryRemoveAiDeclarations(packet, removalOptions, location, changes, out byte[] cleaned)) {
                    int prefixLength = packetOffset - (offset + 8);
                    byte[] rewrittenPayload = new byte[prefixLength + cleaned.Length];
                    Buffer.BlockCopy(data, offset + 8, rewrittenPayload, 0, prefixLength);
                    Buffer.BlockCopy(cleaned, 0, rewrittenPayload, prefixLength, cleaned.Length);
                    WriteChunk(output, "iTXt", rewrittenPayload);
                    reserialized = true;
                } else {
                    output?.Write(data, offset, total);
                }
            } else {
                output?.Write(data, offset, total);
            }
            string type = System.Text.Encoding.ASCII.GetString(data, offset + 4, 4);
            offset += total;
            if (type == "IEND") { foundEnd = true; break; }
        }
        if (!foundEnd) throw new InvalidDataException("PNG does not contain an IEND chunk.");
        if (offset < data.Length) output?.Write(data, offset, data.Length - offset);
        return reserialized;
    }

    private static bool TryGetXmpPacket(byte[] data, int payloadOffset, int payloadLength, out int packetOffset, out int packetLength) {
        packetOffset = packetLength = 0;
        const string keyword = "XML:com.adobe.xmp";
        if (payloadLength < keyword.Length + 5 || !OfficeProvenanceBinary.MatchesAscii(data, payloadOffset, keyword) ||
            data[payloadOffset + keyword.Length] != 0) return false;
        int cursor = payloadOffset + keyword.Length + 1;
        int end = payloadOffset + payloadLength;
        if (cursor + 2 > end || data[cursor++] != 0 || data[cursor++] != 0) return false;
        int terminator = Array.IndexOf(data, (byte)0, cursor, end - cursor);
        if (terminator < 0) return false;
        cursor = terminator + 1;
        terminator = Array.IndexOf(data, (byte)0, cursor, end - cursor);
        if (terminator < 0) return false;
        cursor = terminator + 1;
        packetOffset = cursor;
        packetLength = end - cursor;
        return packetLength > 0;
    }

    private static void WriteChunk(Stream output, string type, byte[] payload) {
        byte[] header = new byte[8];
        OfficeProvenanceBinary.WriteUInt32(header, 0, (uint)payload.Length, littleEndian: false);
        System.Text.Encoding.ASCII.GetBytes(type).CopyTo(header, 4);
        output.Write(header, 0, header.Length);
        output.Write(payload, 0, payload.Length);
        uint crc = ComputeCrc(header, 4, 4, payload);
        byte[] trailer = new byte[4];
        OfficeProvenanceBinary.WriteUInt32(trailer, 0, crc, littleEndian: false);
        output.Write(trailer, 0, trailer.Length);
    }

    private static uint ComputeCrc(byte[] header, int offset, int count, byte[] payload) {
        uint crc = 0xFFFFFFFF;
        for (int index = 0; index < count; index++) crc = UpdateCrc(crc, header[offset + index]);
        for (int index = 0; index < payload.Length; index++) crc = UpdateCrc(crc, payload[index]);
        return crc ^ 0xFFFFFFFF;
    }

    private static bool HasValidCrc(byte[] data, int chunkOffset, int payloadLength) {
        uint expected = OfficeProvenanceBinary.ReadUInt32(data, chunkOffset + 8 + payloadLength, littleEndian: false);
        uint crc = 0xFFFFFFFF;
        for (int index = chunkOffset + 4; index < chunkOffset + 8 + payloadLength; index++) crc = UpdateCrc(crc, data[index]);
        return (crc ^ 0xFFFFFFFF) == expected;
    }

    private static uint UpdateCrc(uint crc, byte value) {
        crc ^= value;
        for (int bit = 0; bit < 8; bit++) crc = (crc & 1) != 0 ? 0xEDB88320U ^ (crc >> 1) : crc >> 1;
        return crc;
    }
}
