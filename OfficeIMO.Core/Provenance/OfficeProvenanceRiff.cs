using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Provenance;

internal static class OfficeProvenanceRiff {
    internal static void Inspect(byte[] data, OfficeProvenanceOptions options, OfficeProvenanceContext context) {
        Walk(data, options, context, output: null, removalOptions: null, changes: null);
    }

    internal static byte[] Remove(byte[] data, OfficeProvenanceRemovalOptions options, List<OfficeProvenanceChange> changes, out bool reserialized) {
        reserialized = false;
        if (!options.RemoveC2paManifests && !options.RemoveAiSourceMetadata) return (byte[])data.Clone();
        using var output = new MemoryStream(data.Length);
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
        if (declaredEndValue < 12 || declaredEndValue > data.Length) throw new InvalidDataException("RIFF size exceeds the asset bounds.");
        int declaredEnd = (int)declaredEndValue;
        int offset = 12;
        int chunkCount = 0;
        while (offset < declaredEnd) {
            if (++chunkCount > options.MaxContainerEntries) {
                throw new InvalidDataException("WebP exceeds the configured container entry limit.");
            }
            if (declaredEnd - offset < 8) throw new InvalidDataException("RIFF contains a truncated chunk header.");
            uint payloadValue = OfficeProvenanceBinary.ReadUInt32(data, offset + 4, littleEndian: true);
            if (payloadValue > int.MaxValue) throw new InvalidDataException("RIFF chunk exceeds the supported size.");
            int payloadLength = (int)payloadValue;
            long totalValue = 8L + payloadLength + (payloadLength & 1);
            if (totalValue > declaredEnd - offset) throw new InvalidDataException("RIFF chunk exceeds the declared container bounds.");
            int total = (int)totalValue;
            bool isC2pa = OfficeProvenanceBinary.MatchesAscii(data, offset, "C2PA");
            if (isC2pa) {
                if (payloadLength > options.MaxManifestBytes) throw new InvalidDataException("RIFF provenance chunk exceeds the configured manifest limit.");
                bool isLast = offset + total == declaredEnd;
                bool valid = isLast && OfficeC2paManifestStore.IsValid(
                    data, offset + 8, payloadLength, options.MaxManifestBytes, options.MaxContainerEntries, out _);
                string location = $"RIFF/C2PA@{offset}";
                context?.Add(new OfficeProvenanceEvidence(OfficeProvenanceCarrierKind.C2paManifest, location, valid, payloadLength));
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
            } else if (OfficeProvenanceBinary.MatchesAscii(data, offset, "XMP ")) {
                byte[] packet = new byte[payloadLength];
                Buffer.BlockCopy(data, offset + 8, packet, 0, payloadLength);
                string location = $"WebP/XMP@{offset}";
                if (context != null) OfficeProvenanceXmp.Inspect(packet, options, context, location);
                if (output != null && removalOptions != null && changes != null &&
                    OfficeProvenanceXmp.TryRemoveAiDeclarations(packet, removalOptions, location, changes, out byte[] cleaned)) {
                    WriteChunk(output, "XMP ", cleaned);
                    reserialized = true;
                } else {
                    output?.Write(data, offset, total);
                }
            } else {
                output?.Write(data, offset, total);
            }
            offset += total;
        }
        if (offset != declaredEnd) throw new InvalidDataException("RIFF chunks do not end on the declared boundary.");
        if (declaredEnd < data.Length) output?.Write(data, declaredEnd, data.Length - declaredEnd);
        return reserialized;
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
